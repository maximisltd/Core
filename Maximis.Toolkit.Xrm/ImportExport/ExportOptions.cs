using Maximis.Toolkit.Xrm.EntitySerialisation;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    public class ExportOptions : DisplayStringOptions
    {
        public int ExportLimit { get; set; }

        public QueryExpression QueryExpression { get; set; }

        public int RecordsPerPage { get; set; }

        public List<EntitySerialiserScope> Scopes { get; set; }
    }
}