namespace Maximis.Toolkit.Xrm.ImportExport
{
    public class CsvExportOptions : ExportOptions
    {
        public CsvExportOptions()
            : base()
        {
        }

        public CsvExportOptions(bool forImport)
            : base()
        {
            if (forImport)
            {
                this.DateFormat = "yyyy-MM-dd HH:mm:ss";
                this.HeaderFormat = "{0} ({1})";
                this.LookupFormat = "[{0}] [{1}] [{2}]";
                this.OptionSetFormat = "[{0}] [{1}]";
            }
        }

        public string HeaderFormat { get; set; }
    }
}