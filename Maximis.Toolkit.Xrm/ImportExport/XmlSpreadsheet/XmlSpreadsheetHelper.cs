using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;

namespace Maximis.Toolkit.Xrm.ImportExport.XmlSpreadsheet
{
    public static class XmlSpreadsheetHelper
    {
        private static DisplayStringOptions formats = new DisplayStringOptions
        {
            DateFormat = "yyyy-MM-ddTHH:mm:ss.fff",
            BoolTrue = "1",
            BoolFalse = "0"
        };

        public delegate DataRow ModifyDataRow(DataRow dataRow, string currentColumn, Entity entity);

        /// <summary>
        /// Return the results of a QueryExpression as an XmlSpreadsheet
        /// </summary>
        public static XmlDocument GenerateXmlSpreadsheet(CrmContext context, QueryExpression query, string worksheetName)
        {
            return GenerateXmlSpreadsheet(worksheetName, RetrieveData(context, query));
        }

        /// <summary>
        /// Constructs an XmlSpreadsheet using single row of data (presented vertically instead of horizontally)
        /// </summary>
        public static XmlDocument GenerateXmlSpreadsheet(string worksheetName, DataRow singleRow)
        {
            // Ensure we have data
            bool hasData = (singleRow != null && singleRow.DataItems.Count > 0);

            // Load the Template spreadsheet
            XmlSpreadsheet xs = GetEmptySpreadsheet(worksheetName, !hasData);

            if (hasData)
            {
                // Loop through Data Items
                foreach (DataItem di in singleRow.DataItems)
                {
                    // Add a Row tag
                    XmlElement row = AppendElement(xs, xs.WorksheetTable, "Row");

                    // Add a Label Cell
                    AppendCellElement(xs, row, di.Label, di.LabelStyleId);

                    // Add a Value Cell
                    AppendCellElement(xs, row, di);
                }

                // Return the XML Document
                return xs.XmlDocument;
            }

            return null;
        }

        /// <summary>
        /// Convert a List of DataRow into an XML Spreadsheet.
        /// </summary>
        public static XmlDocument GenerateXmlSpreadsheet(string worksheetName, List<DataRow> sourceData,
            Func<DataRow, bool> includeRow = null, params string[] cols)
        {
            // Ensure we have data
            bool hasData = (sourceData != null && sourceData.Count > 0);

            // Load the Template spreadsheet
            XmlSpreadsheet xs = GetEmptySpreadsheet(worksheetName, !hasData);

            if (hasData)
            {
                // Flag to determine if we need to add Header values
                bool isFirstRow = true;
                XmlElement headerRow = AppendElement(xs, xs.WorksheetTable, "Row");

                // Loop through Data
                foreach (DataRow dr in sourceData)
                {
                    // Check Row should be included
                    if (includeRow != null && !includeRow(dr)) continue;

                    // Add a Row tag
                    XmlElement row = AppendElement(xs, xs.WorksheetTable, "Row");

                    // Get the DataItems to loop through
                    IEnumerable<DataItem> dataItems = (cols == null || cols.Length == 0)
                        ? dr.DataItems
                        : dr.DataItems.Where(q => cols.Contains(q.Key));

                    // Loop through the DataItems
                    foreach (DataItem di in dataItems)
                    {
                        // Add Header Cell if necessary
                        if (isFirstRow)
                        {
                            AppendCellElement(xs, headerRow, di.Label, di.LabelStyleId);
                        }

                        // Add Data Cell
                        AppendCellElement(xs, row, di);
                    }

                    // Don't add headers on subsequent rows
                    isFirstRow = false;
                }

                // Return the XML Document
                return xs.XmlDocument;
            }

            return null;
        }

        public static DataRow GetDataRowFromEntity(CrmContext context, Entity entity, IEnumerable<string> columns, ModifyDataRow modifyDataRow = null)
        {
            DataRow dataRow = new DataRow();

            EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, entity.LogicalName);

            // Loop through each Attribute
            foreach (string column in columns)
            {
                AttributeMetadata attMeta = MetadataHelper.GetAttributeMetadata(entityMeta, column);

                // Create a DataItem
                DataItem dataItem = new DataItem
                {
                    Key = column,
                    Label = attMeta.DisplayName.UserLocalizedLabel.Label,
                    LabelStyleId = "bold"
                };

                // If Attribute has value, add to DataItem (with String version)
                if (entity.HasAttributeWithValue(column))
                {
                    dataItem.Value = entity[column];
                    dataItem.ValueString = MetadataHelper.GetAttributeValueAsDisplayString(context, entity, column, formats);
                }

                // Add DataItem to DataRow
                dataRow.DataItems.Add(dataItem);

                // Modify Data Row if a delegate is defined
                if (modifyDataRow != null)
                {
                    dataRow = modifyDataRow(dataRow, column, entity);
                }
            }

            return dataRow;
        }

        /// <summary>
        /// Retrieves data from CRM and converts it to a List of DataRow objects
        /// </summary>
        public static List<DataRow> RetrieveData(CrmContext context, QueryExpression query, ModifyDataRow modifyDataRow = null)
        {
            // Create the object to return
            List<DataRow> allData = new List<DataRow>();

            // Loop through all retrieved entities
            EntityCollection results = null;
            while (QueryHelper.RetrieveEntitiesWithPaging(context.OrganizationService, query, ref results))
            {
                foreach (Entity entity in results.Entities)
                {
                    // Create a new DataRow
                    DataRow dataRow = GetDataRowFromEntity(context, entity, query.ColumnSet.Columns, modifyDataRow);

                    // Add the DataRow to the allData collection
                    allData.Add(dataRow);
                }
            }

            // Return
            return allData;
        }

        /// <summary>
        /// Creates a Cell element containing the value of the DataItem supplied.
        /// </summary>
        private static XmlElement AppendCellElement(XmlSpreadsheet xs, XmlElement row, DataItem di,
            string styleId = null)
        {
            XmlElement cell = AppendElement(xs, row, "Cell");

            if (di.Value != null)
            {
                // If DataItem has a Value, create a Data tag
                XmlElement data = AppendElement(xs, cell, "Data");
                SetSSAttribute(data, "Type", di.CellType);

                // Set the Cell Style
                if (!string.IsNullOrEmpty(styleId))
                {
                    SetSSAttribute(cell, "StyleID", styleId);
                }
                else if (!string.IsNullOrEmpty(di.ValueStyleId))
                {
                    SetSSAttribute(cell, "StyleID", di.ValueStyleId);
                }

                // Populate Data tag
                data.InnerText = di.ValueString;
            }

            return cell;
        }

        /// <summary>
        /// Creates a Cell element containing the supplied string
        /// </summary>
        private static XmlElement AppendCellElement(XmlSpreadsheet xs, XmlElement row, string text,
            string styleId = null)
        {
            XmlElement stringCell = AppendElement(xs, row, "Cell");
            XmlElement stringData = AppendElement(xs, stringCell, "Data");
            SetSSAttribute(stringData, "Type", "String");
            if (!string.IsNullOrEmpty(styleId))
            {
                SetSSAttribute(stringCell, "StyleID", styleId);
            }
            stringData.InnerText = text;
            return stringCell;
        }

        /// <summary>
        /// Wrapper method to append an XML Element to a Parent
        /// </summary>
        private static XmlElement AppendElement(XmlSpreadsheet xs, XmlElement appendTo, string name)
        {
            return AppendElement(xs.XmlDocument, appendTo, name);
        }

        /// <summary>
        /// Wrapper method to append an XML Element to a Parent
        /// </summary>
        private static XmlElement AppendElement(XmlDocument xd, XmlElement appendTo, string name)
        {
            XmlElement result = xd.CreateElement(name, xd.DocumentElement.NamespaceURI);
            appendTo.AppendChild(result);
            return result;
        }

        /// <summary>
        /// Retrieves the EmptySpreadsheet Embedded Resource
        /// </summary>
        private static XmlSpreadsheet GetEmptySpreadsheet(string worksheetName, bool addNoDataError = false)
        {
            // Retrieve Empty Spreadsheet from Assembly
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resName = assembly.GetManifestResourceNames()
                .SingleOrDefault(q => q.EndsWith("EmptySpreadsheet.xml"));
            XmlDocument xd = new XmlDocument();
            using (Stream s = assembly.GetManifestResourceStream(resName))
            {
                xd.Load(s);
            }

            // Add Worksheet
            XmlElement worksheet = AppendElement(xd, xd.DocumentElement, "Worksheet");
            SetSSAttribute(worksheet, "Name", worksheetName);
            XmlElement table = AppendElement(xd, worksheet, "Table");
            SetSSAttribute(table, "DefaultColumnWidth", "120");

            // Add Error if required
            if (addNoDataError)
            {
                XmlElement errorRow = AppendElement(xd, table, "Row");
                XmlElement errorCell = AppendElement(xd, errorRow, "Cell");
                XmlElement errorData = AppendElement(xd, errorCell, "Data");
                SetSSAttribute(errorData, "Type", "Error");
                errorData.InnerText = "There are no records available for this report.";
            }

            // Return
            return new XmlSpreadsheet { XmlDocument = xd, WorksheetTable = table };
        }

        /// <summary>
        /// Wrapper method to add an attribute in the "ss:" namespace to an Element
        /// </summary>
        private static void SetSSAttribute(XmlElement xmlElement, string name, string val)
        {
            XmlAttribute attribute = xmlElement.OwnerDocument.CreateAttribute("ss", name,
                "urn:schemas-microsoft-com:office:spreadsheet");
            attribute.InnerText = val;
            xmlElement.SetAttributeNode(attribute);
        }

        private class XmlSpreadsheet
        {
            public XmlElement WorksheetTable { get; set; }

            public XmlDocument XmlDocument { get; set; }
        }
    }
}