using Microsoft.Xrm.Sdk.Query;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    public class ExportOptions : DisplayStringOptions
    {
        public int ExportLimit { get; set; }

        public QueryExpression QueryExpression { get; set; }

        public int RecordsPerPage { get; set; }
    }
}