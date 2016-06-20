using Maximis.Toolkit.IO;
using Maximis.Toolkit.Xrm.EntitySerialisation;
using Microsoft.Xrm.Sdk;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    public static class XmlImportExportHelper
    {
        #region Export

        public static void Export(CrmContext context, XmlTextWriter output, ExportOptions options)
        {
            // Get Entity Serialiser
            EntitySerialiser ser = GetEntitySerialiser(context, options);

            // Serialise to XML file
            output.WriteStartElement("allentities");
            EntityCollection results = null;
            while (QueryHelper.RetrieveEntitiesWithPaging(context.OrganizationService, options.QueryExpression, ref results, options.RecordsPerPage))
            {
                foreach (Entity entity in results.Entities)
                {
                    output.WriteRaw(ser.SerialiseEntity(entity));
                }
            }

            // Close off XML file
            output.WriteEndElement();
            output.Flush();
        }

        public static void ExportMultipleFiles(CrmContext context, string folderPath, ExportOptions options)
        {
            FileHelper.EnsureDirectoryExists(folderPath, PathType.Directory);

            EntitySerialiser ser = GetEntitySerialiser(context, options);

            EntityCollection results = null;
            int index = 0;
            while (QueryHelper.RetrieveEntitiesWithPaging(context.OrganizationService, options.QueryExpression, ref results, options.RecordsPerPage))
            {
                foreach (Entity entity in results.Entities)
                {
                    Trace.WriteLine(string.Format("Exporting '{0}' with id '{1:N}'", entity.LogicalName, entity.Id));

                    string filePath = Path.Combine(folderPath, string.Format("{0:000000}_{1}_{2:N}.xml", index++, entity.LogicalName, entity.Id));
                    using (StreamWriter sw = new StreamWriter(filePath))
                    using (XmlTextWriter xtw = new XmlTextWriter(sw))
                    {
                        xtw.WriteRaw(ser.SerialiseEntity(entity));
                    }
                }
            }
        }

        public static void ExportSingleFile(CrmContext context, string filePath, ExportOptions options)
        {
            FileHelper.EnsureDirectoryExists(filePath, PathType.File);
            using (StreamWriter sw = new StreamWriter(filePath))
            using (XmlTextWriter xtw = new XmlTextWriter(sw))
            {
                Export(context, xtw, options);
            }
        }

        private static EntitySerialiser GetEntitySerialiser(CrmContext context, ExportOptions options)
        {
            if (options.Scopes == null || options.Scopes.Count == 0)
            {
                EntitySerialiserScope scope = new EntitySerialiserScope { EntityType = options.QueryExpression.EntityName, Columns = ImportExportHelper.GetAllQueryAttributes(context, options.QueryExpression) };
                return new EntitySerialiser(context, scope);
            }
            else
            {
                return new EntitySerialiser(context, options.Scopes);
            }
        }

        #endregion Export

        #region Import

        public static void ImportMultipleFiles(CrmContext context, string folderPath, ImportOptions options)
        {
            using (XmlImportManager xmlImport = new XmlImportManager(context, options))
            {
                foreach (string fileName in Directory.EnumerateFiles(folderPath, "*.xml", SearchOption.AllDirectories))
                {
                    xmlImport.AddForImport(File.ReadAllText(fileName));
                }
            }
        }

        public static void ImportSingleFile(CrmContext context, XmlTextReader input, ImportOptions options)
        {
            using (XmlImportManager xmlImport = new XmlImportManager(context, options))
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

        public static void ImportSingleFile(CrmContext context, string filePath, ImportOptions options)
        {
            using (XmlTextReader xr = new XmlTextReader(File.OpenText(filePath)))
            {
                ImportSingleFile(context, xr, options);
            }
        }

        #endregion Import
    }
}