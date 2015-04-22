using Maximis.Toolkit.Csv;
using Maximis.Toolkit.IO;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    public static class CsvImportExportHelper
    {
        #region Export

        public static void Export(IOrganizationService orgService, string filePath, CsvExportOptions options)
        {
            FileHelper.EnsureDirectoryExists(filePath, PathType.File);
            using (StreamWriter sw = new StreamWriter(filePath)) Export(orgService, sw, options);
        }

        public static void Export(IOrganizationService orgService, TextWriter output, CsvExportOptions options)
        {
            MetadataCache metaCache = new MetadataCache();

            string[] attributes = ImportExportHelper.GetAllQueryAttributes(orgService, options.QueryExpression, metaCache);

            string[] columnHeadings = string.IsNullOrEmpty(options.HeaderFormat) ? attributes :
               GetColumnHeadings(orgService, options.QueryExpression.EntityName, attributes, metaCache, options.HeaderFormat);

            CsvWriter csvWriter = new CsvWriter(output, columnHeadings);

            EntityCollection results = null;
            int exportCount = 0;
            while (QueryHelper.RetrieveEntitiesWithPaging(orgService, options.QueryExpression, ref results, options.RecordsPerPage))
            {
                foreach (Entity entity in results.Entities)
                {
                    Trace.WriteLine(string.Format("Exporting '{0}' with id '{1:N}'", entity.LogicalName, entity.Id));

                    for (int i = 0; i < attributes.Length; i++)
                    {
                        string val = MetadataHelper.GetAttributeValueAsDisplayString(orgService, metaCache, entity, attributes[i], options);
                        csvWriter.WriteValue(i, val);
                    }
                }
                if (++exportCount < options.ExportLimit) return;
            }
        }

        private static string[] GetColumnHeadings(IOrganizationService orgService, string entityName, string[] attributes, MetadataCache metaCache, string headerFormat)
        {
            EntityMetadata meta = metaCache.GetEntityMetadata(orgService, entityName);
            return attributes.Select(q => string.Format(headerFormat, q, GetLabel(meta.Attributes.SingleOrDefault(x => x.LogicalName == q.RightOfLast('.'))))).ToArray();
        }

        private static string GetLabel(AttributeMetadata attrMeta)
        {
            if (attrMeta == null || attrMeta.DisplayName == null || attrMeta.DisplayName.UserLocalizedLabel == null || string.IsNullOrEmpty(attrMeta.DisplayName.UserLocalizedLabel.Label))
                return "[No Display Name]";
            return attrMeta.DisplayName.UserLocalizedLabel.Label;
        }

        #endregion Export

        #region Import

        private static readonly Regex REGEX_SQUAREBRACKETS = new Regex(@"\[.*?\]", RegexOptions.Compiled);

        public static void Import(IOrganizationService orgService, string filePath, CsvImportOptions options)
        {
            using (StreamReader sr = new StreamReader(filePath))
            {
                Import(orgService, sr, options);
            }
        }

        public static void Import(IOrganizationService orgService, TextReader input, CsvImportOptions options)
        {
            // Get Entity Metadata
            MetadataCache metaCache = new MetadataCache();
            EntityMetadata meta = metaCache.GetEntityMetadata(orgService, options.EntityType);

            using (XmlImportManager xmlImport = new XmlImportManager(orgService, metaCache, options))
            using (CsvReader csvParser = new CsvReader(input))
            {
                // Find the column containing the ID (there may not be one)
                int idIndex = csvParser.HeadingIndex.ContainsKey(meta.PrimaryIdAttribute) ? csvParser.HeadingIndex[meta.PrimaryIdAttribute] : int.MinValue;

                // Loop through CSV
                while (!csvParser.EndOfData)
                {
                    string[] csvData = null;
                    try { csvData = csvParser.ReadFields(); }
                    catch (MalformedLineException)
                    {
                        Trace.WriteLine("==========");
                        Trace.WriteLine(csvParser.ErrorLine);
                        Trace.WriteLine("==========");
                        throw;
                    }

                    // Format Data as Serialised Entity XML
                    StringBuilder sb = new StringBuilder();
                    using (StringWriter sw = new StringWriter(sb))
                    using (XmlTextWriter xtw = new XmlTextWriter(sw))
                    {
                        xtw.WriteStartElement("ent");
                        if (idIndex >= 0) xtw.WriteAttributeString("id", csvData[idIndex]);
                        xtw.WriteAttributeString("ln", meta.LogicalName);

                        foreach (string attributeName in csvParser.HeadingIndex.Keys.Where(q => q != meta.PrimaryIdAttribute))
                        {
                            int fieldIndex = csvParser.HeadingIndex[attributeName];
                            if (csvData.Length <= fieldIndex) break;

                            string val = csvData[fieldIndex];
                            if (string.IsNullOrWhiteSpace(val))
                            {
                                if (options.SetToNullIfEmpty != null && options.SetToNullIfEmpty.Contains(attributeName))
                                {
                                    xtw.WriteStartElement("attr");
                                    xtw.WriteAttributeString("ln", attributeName);
                                    xtw.WriteEndElement();
                                }
                                continue;
                            }

                            bool writeString = true;
                            AttributeMetadata attrMeta = meta.Attributes.SingleOrDefault(q => q.LogicalName == attributeName);

                            xtw.WriteStartElement("attr");
                            xtw.WriteAttributeString("ln", attributeName);

                            if (attrMeta != null)
                            {
                                switch (attrMeta.AttributeType)
                                {
                                    case AttributeTypeCode.Lookup:
                                    case AttributeTypeCode.Owner:
                                    case AttributeTypeCode.Customer:
                                        EntityReference entityRef = GetEntityReference(metaCache, orgService, val);
                                        if (entityRef != null)
                                        {
                                            writeString = false;
                                            xtw.WriteStartElement("ent");
                                            xtw.WriteAttributeString("ln", entityRef.LogicalName);
                                            xtw.WriteAttributeString("id", entityRef.Id.ToString("D"));
                                            xtw.WriteStartElement("attr");
                                            xtw.WriteAttributeString("ln", metaCache.GetEntityMetadata(orgService, entityRef.LogicalName).PrimaryNameAttribute);
                                            xtw.WriteString(entityRef.Name);
                                            xtw.WriteEndElement();
                                            xtw.WriteEndElement();
                                        }
                                        break;

                                    case AttributeTypeCode.PartyList:
                                        throw new Exception("Party List not supported!");

                                    case AttributeTypeCode.Picklist:
                                    case AttributeTypeCode.State:
                                    case AttributeTypeCode.Status:
                                        EnumAttributeMetadata enumMeta = (EnumAttributeMetadata)attrMeta;
                                        foreach (string valItem in GetValuesFromSquareBrackets(val))
                                        {
                                            int optionVal = int.MinValue;
                                            if (int.TryParse(valItem, out optionVal))
                                            {
                                                OptionMetadata option = enumMeta.OptionSet.Options.SingleOrDefault(q => q.Value.Value == optionVal);
                                                if (option != null)
                                                {
                                                    writeString = false;
                                                    xtw.WriteAttributeString("val", option.Value.Value.ToString());
                                                    xtw.WriteString(option.Label.UserLocalizedLabel.Label);
                                                }
                                            }
                                        }
                                        break;
                                }
                            }
                            if (writeString) xtw.WriteString(val);
                            xtw.WriteEndElement();
                        }
                        xtw.WriteEndElement();

                        xmlImport.AddForImport(sb.ToString());
                    }
                }
            }
        }

        private static EntityReference GetEntityReference(MetadataCache metaCache, IOrganizationService orgService, string val)
        {
            EntityReference result = new EntityReference();
            foreach (string s in GetValuesFromSquareBrackets(val))
            {
                if (result.Id == Guid.Empty)
                {
                    Guid id = Guid.Empty;
                    if (Guid.TryParse(s, out id))
                    {
                        result.Id = id;
                        continue;
                    }
                }
                if (string.IsNullOrEmpty(result.LogicalName))
                {
                    try
                    {
                        EntityMetadata meta = metaCache.GetEntityMetadata(orgService, s);
                        result.LogicalName = meta.LogicalName;
                        continue;
                    }
                    catch { }
                }
                result.Name = s;
            }

            if (result.Id == Guid.Empty || string.IsNullOrEmpty(result.LogicalName)) return null;
            return result;
        }

        private static string[] GetValuesFromSquareBrackets(string val)
        {
            List<string> result = new List<string>();

            foreach (Match m in REGEX_SQUAREBRACKETS.Matches(val))
            {
                result.Add(m.Value.Substring(1, m.Value.Length - 2));
                val = val.Replace(m.Value, string.Empty);
            }
            if (!string.IsNullOrWhiteSpace(val))
            {
                result.Add(val.Trim());
            }
            return result.ToArray();
        }

        #endregion Import
    }
}