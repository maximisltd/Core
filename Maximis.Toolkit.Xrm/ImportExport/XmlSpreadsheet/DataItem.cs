using System;

namespace Maximis.Toolkit.Xrm.ImportExport.XmlSpreadsheet
{
    public class DataItem
    {
        private string cellType;
        private string valueString;

        public string CellType
        {
            get
            {
                if (string.IsNullOrEmpty(cellType))
                {
                    PrepareForXmlSpreadsheet();
                }
                return cellType;
            }
            set { cellType = value; }
        }

        public string Key { get; set; }

        public string Label { get; set; }

        public string LabelStyleId { get; set; }

        public object Value { get; set; }

        public string ValueString
        {
            get
            {
                if (string.IsNullOrEmpty(valueString))
                {
                    PrepareForXmlSpreadsheet();
                }
                return valueString;
            }
            set { valueString = value; }
        }

        public string ValueStyleId { get; set; }

        private void PrepareForXmlSpreadsheet()
        {
            if (Value == null)
            {
                ValueString = null;
                CellType = null;
            }
            else
            {
                Type t = Value.GetType();
                switch (t.Name)
                {
                    case "Byte":
                    case "SByte":
                    case "Int16":
                    case "UInt16":
                    case "Int32":
                    case "UInt32":
                    case "Int64":
                    case "UInt64":
                    case "Single":
                    case "Double":
                    case "Decimal":
                        if (string.IsNullOrEmpty(valueString)) valueString = Value.ToString();
                        if (string.IsNullOrEmpty(cellType)) cellType = "Number";
                        break;

                    case "Boolean":
                        if (string.IsNullOrEmpty(valueString)) valueString = (bool)Value ? "1" : "0";
                        if (string.IsNullOrEmpty(cellType)) cellType = "Boolean";
                        break;

                    case "DateTime":
                        if (string.IsNullOrEmpty(valueString))
                            valueString = ((DateTime)Value).ToString("yyyy-MM-ddTHH:mm:ss.fff");
                        if (string.IsNullOrEmpty(cellType)) cellType = "DateTime";
                        ValueStyleId = "dateTime";
                        break;

                    default:
                        if (string.IsNullOrEmpty(valueString)) valueString = Value.ToString();
                        if (string.IsNullOrEmpty(cellType)) cellType = "String";
                        break;
                }
            }
        }
    }
}