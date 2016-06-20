using Maximis.Toolkit.Csv;
using Maximis.Toolkit.IO;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
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

        public static void Export(CrmContext context, string filePath, CsvExportOptions options)
        {
            FileHelper.EnsureDirectoryExists(filePath, PathType.File);
            using (StreamWriter sw = new StreamWriter(filePath)) Export(context, sw, options);
        }

        public static void Export(CrmContext context, TextWriter output, CsvExportOptions options)
        {
            string[] attributes = ImportExportHelper.GetAllQueryAttributes(context, options.QueryExpression);
            if (attributes.Length == 0) throw new InvalidOperationException("Query ColumnSet is missing or empty.");

            string[] columnHeadings = string.IsNullOrEmpty(options.HeaderFormat) ? attributes :
               GetColumnHeadings(context, options.QueryExpression.EntityName, attributes, options.HeaderFormat);

            CsvWriter csvWriter = new CsvWriter(output, true, columnHeadings);

            EntityCollection results = null;
            int exportCount = 0;
            while (QueryHelper.RetrieveEntitiesWithPaging(context.OrganizationService, options.QueryExpression, ref results, options.RecordsPerPage))
            {
                foreach (Entity entity in results.Entities)
                {
                    Trace.WriteLine(string.Format("Exporting '{0}' with id '{1:N}'", entity.LogicalName, entity.Id));

                    for (int i = 0; i < attributes.Length; i++)
                    {
                        string val = MetadataHelper.GetAttributeValueAsDisplayString(context, entity, attributes[i], options);
                        csvWriter.WriteValue(i, val);
                    }
                }
                if (++exportCount < options.ExportLimit) return;
            }
        }

        public static void ExportRelationships(CrmContext context, string filePath, params string[] entityTypes)
        {
            FileHelper.EnsureDirectoryExists(filePath, PathType.File);
            using (StreamWriter sw = new StreamWriter(filePath)) ExportRelationships(context, sw, entityTypes);
        }

        public static void ExportRelationships(CrmContext context, TextWriter output, params string[] entityTypes)
        {
            List<Guid> processed = new List<Guid>();

            CsvWriter csvWriter = new CsvWriter(output, true, new string[] { "Record 1 Type", "Record 1 Id", "Record 2 Type", "Record 2 Id", "Relationship" });

            EntityCollection results = null;

            foreach (string type1 in entityTypes)
            {
                results = null;
                EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, type1, EntityFilters.Entity | EntityFilters.Relationships);
                while (QueryHelper.RetrieveEntitiesWithPaging(context.OrganizationService, new QueryExpression(entityMeta.LogicalName), ref results))
                {
                    foreach (Entity entity in results.Entities)
                    {
                        foreach (string type2 in entityTypes.Where(q => q != type1))
                        {
                            foreach (OneToManyRelationshipMetadata relMeta in entityMeta.OneToManyRelationships.Where(q => q.ReferencingEntity == type2 || q.ReferencedEntity == type2))
                            {
                                ExportRelationshipsWorker(context, entity.ToEntityReference(), relMeta, csvWriter, ref processed);
                            }
                            foreach (OneToManyRelationshipMetadata relMeta in entityMeta.ManyToOneRelationships.Where(q => q.ReferencingEntity == type2 || q.ReferencedEntity == type2))
                            {
                                ExportRelationshipsWorker(context, entity.ToEntityReference(), relMeta, csvWriter, ref processed);
                            }
                            foreach (ManyToManyRelationshipMetadata relMeta in entityMeta.ManyToManyRelationships.Where(q => q.Entity1LogicalName == type2 || q.Entity2LogicalName == type2))
                            {
                                ExportRelationshipsWorker(context, entity.ToEntityReference(), relMeta, csvWriter, ref processed);
                            }
                        }
                    }
                }
            }
        }

        private static void ExportRelationshipsWorker(CrmContext context, EntityReference entityReference, RelationshipMetadataBase relMeta, CsvWriter csvWriter, ref List<Guid> processed)
        {
            foreach (Entity related in QueryHelper.RetrieveRelatedEntities(context.OrganizationService, entityReference, relMeta).Entities)
            {
                if (processed.Contains(related.Id)) continue;
                csvWriter.WriteValue(0, entityReference.LogicalName);
                csvWriter.WriteValue(1, entityReference.Id.ToString("N"));
                csvWriter.WriteValue(2, related.LogicalName);
                csvWriter.WriteValue(3, related.Id.ToString("N"));
                csvWriter.WriteValue(4, relMeta.SchemaName);
            }
            processed.Add(entityReference.Id);
        }

        private static string[] GetColumnHeadings(CrmContext context, string entityName, string[] attributes, string headerFormat)
        {
            EntityMetadata meta = MetadataHelper.GetEntityMetadata(context, entityName);
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

        public static void Import(CrmContext context, string filePath, CsvImportOptions options)
        {
            using (StreamReader sr = new StreamReader(filePath))
            {
                Import(context, sr, options);
            }
        }

        public static void Import(CrmContext context, TextReader input, CsvImportOptions options)
        {
            // Get Entity Metadata
            EntityMetadata meta = MetadataHelper.GetEntityMetadata(context, options.EntityType);

            using (XmlImportManager xmlImport = new XmlImportManager(context, options))
            using (CsvReader csvParser = new CsvReader(input))
            {
                // If Headings are in format "attributename (Attribute Label)", replace with
                // "attributename" only
                csvParser.HeadingIndex = csvParser.HeadingIndex.ToDictionary(k => k.Key.LeftOfFirst('(').Trim(), v => v.Value);

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
                                        LookupAttributeMetadata lookupMeta = (LookupAttributeMetadata)attrMeta;
                                        EntityReference entityRef = GetEntityReference(context, val, lookupMeta.Targets);
                                        if (entityRef != null)
                                        {
                                            writeString = false;
                                            xtw.WriteStartElement("ent");
                                            xtw.WriteAttributeString("ln", entityRef.LogicalName);
                                            xtw.WriteAttributeString("id", entityRef.Id.ToString("D"));
                                            xtw.WriteStartElement("attr");
                                            xtw.WriteAttributeString("ln", MetadataHelper.GetEntityMetadata(context, entityRef.LogicalName).PrimaryNameAttribute);
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

        public static void ImportRelationships(CrmContext context, string filePath)
        {
            using (StreamReader sr = new StreamReader(filePath))
            {
                ImportRelationships(context, sr);
            }
        }

        public static void ImportRelationships(CrmContext context, StreamReader input)
        {
            using (CsvReader csvParser = new CsvReader(input))
            {
                while (!csvParser.EndOfData)
                {
                    string[] csvData = csvParser.ReadFields();
                    EntityReference relateFrom = new EntityReference(csvData[0], new Guid(csvData[1]));
                    EntityReference relateTo = new EntityReference(csvData[2], new Guid(csvData[3]));
                    UpdateHelper.RelateEntities(context.OrganizationService, csvData[4], relateFrom, relateTo);
                }
            }
        }

        private static EntityReference GetEntityReference(CrmContext context, string val, string[] lookupTargets)
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
                        EntityMetadata meta = MetadataHelper.GetEntityMetadata(context, s);
                        result.LogicalName = meta.LogicalName;
                        continue;
                    }
                    catch { }
                }
                result.Name = s;
            }

            if (result.Id == Guid.Empty) return null;

            if (string.IsNullOrEmpty(result.LogicalName))
            {
                if (lookupTargets.Length == 1)
                {
                    result.LogicalName = lookupTargets[0];
                }
                else return null;
            }

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