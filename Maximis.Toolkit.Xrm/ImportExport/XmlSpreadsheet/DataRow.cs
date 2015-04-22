using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm.ImportExport.XmlSpreadsheet
{
    public class DataRow
    {
        public DataRow()
        {
            DataItems = new List<DataItem>();
        }

        public List<DataItem> DataItems { get; set; }
    }
}