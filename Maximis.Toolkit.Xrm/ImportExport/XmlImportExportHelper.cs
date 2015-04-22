using Maximis.Toolkit.IO;
using Maximis.Toolkit.Xrm.EntitySerialisation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    public static class XmlImportExportHelper
    {
        #region Export

        public static void Export(IOrganizationService orgService, XmlTextWriter output, ExportOptions options)
        {
            // Get Entity Serialiser
            EntitySerialiser ser = GetEntitySerialiser(orgService, options.QueryExpression);

            // Serialise to XML file
            output.WriteStartElement("allentities");
            EntityCollection results = null;
            while (QueryHelper.RetrieveEntitiesWithPaging(orgService, options.QueryExpression, ref results, options.RecordsPerPage))
            {
                foreach (Entity entity in results.Entities)
                {
                    output.WriteRaw(ser.SerialiseEntity(orgService, entity));
                }
            }

            // Close off XML file
            output.WriteEndElement();
            output.Flush();
        }

        public static void ExportMultipleFiles(IOrganizationService orgService, string folderPath, ExportOptions options)
        {
            FileHelper.EnsureDirectoryExists(folderPath, PathType.Directory);

            EntitySerialiser ser = GetEntitySerialiser(orgService, options.QueryExpression);

            EntityCollection results = null;
            int index = 0;
            while (QueryHelper.RetrieveEntitiesWithPaging(orgService, options.QueryExpression, ref results, options.RecordsPerPage))
            {
                foreach (Entity entity in results.Entities)
                {
                    Trace.WriteLine(string.Format("Exporting '{0}' with id '{1:N}'", entity.LogicalName, entity.Id));

                    string filePath = Path.Combine(folderPath, string.Format("{0:000000}_{1}_{2:N}.xml", index++, entity.LogicalName, entity.Id));
                    using (StreamWriter sw = new StreamWriter(filePath))
                    using (XmlTextWriter xtw = new XmlTextWriter(sw))
                    {
                        xtw.WriteRaw(ser.SerialiseEntity(orgService, entity));
                    }
                }
            }
        }

        public static void ExportSingleFile(IOrganizationService orgService, string filePath, ExportOptions options)
        {
            FileHelper.EnsureDirectoryExists(filePath, PathType.File);
            using (StreamWriter sw = new StreamWriter(filePath))
            using (XmlTextWriter xtw = new XmlTextWriter(sw))
            {
                Export(orgService, xtw, options);
            }
        }

        private static EntitySerialiser GetEntitySerialiser(IOrganizationService orgService, QueryExpression query)
        {
            MetadataCache metaCache = new MetadataCache();
            EntitySerialiserScope scope = new EntitySerialiserScope { EntityType = query.EntityName, Columns = ImportExportHelper.GetAllQueryAttributes(orgService, query, metaCache) };
            return new EntitySerialiser(metaCache, scope);
        }

        #endregion Export

        #region Import

        public static void ImportMultipleFiles(IOrganizationService orgService, string folderPath, ImportOptions options)
        {
            using (XmlImportManager xmlImport = new XmlImportManager(orgService, new MetadataCache(), options))
            {
                foreach (string fileName in Directory.EnumerateFiles(folderPath, "*.xml"))
                {
                    xmlImport.AddForImport(File.ReadAllText(fileName));
                }
            }
        }

        public static void ImportSingleFile(IOrganizationService orgService, XmlTextReader input, ImportOptions options)
        {
            using (XmlImportManager xmlImport = new XmlImportManager(orgService, new MetadataCache(), options))
            {
                while (input.Read())
                {
                    if (input.IsStartElement() && input.Name == "ent")
                    {
                        xmlImport.AddForImport(input.ReadOuterXml());
                    }
                }
            }
        }

        public static void ImportSingleFile(IOrganizationService orgService, string filePath, ImportOptions options)
        {
            using (XmlTextReader xr = new XmlTextReader(File.OpenText(filePath)))
            {
                ImportSingleFile(orgService, xr, options);
            }
        }

        #endregion Import
    }
}