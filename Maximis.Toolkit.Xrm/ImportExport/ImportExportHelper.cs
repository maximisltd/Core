using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using System.Linq;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    internal static class ImportExportHelper
    {
        internal static string[] GetAllQueryAttributes(IOrganizationService orgService, QueryExpression query, MetadataCache metaCache)
        {
            List<string> attributes = new List<string>();
            attributes.AddRange(GetAllAttributesFromColumnSet(orgService, metaCache, query.EntityName, query.ColumnSet));
            foreach (LinkEntity le in query.LinkEntities)
            {
                attributes.AddRange(GetAllAttributesFromColumnSet(orgService, metaCache, le.LinkToEntityName, le.Columns, le.EntityAlias));
            }
            return attributes.ToArray();
        }

        private static IEnumerable<string> GetAllAttributesFromColumnSet(IOrganizationService orgService, MetadataCache metaCache, string entityType, ColumnSet columnSet, string alias = null)
        {
            if (columnSet.AllColumns)
            {
                List<string> attributes = new List<string>();
                EntityMetadata meta = metaCache.GetEntityMetadata(orgService, entityType);
                if (!string.IsNullOrEmpty(meta.PrimaryIdAttribute)) attributes.Add(GetFullAttributeName(meta.PrimaryIdAttribute, alias));
                if (!string.IsNullOrEmpty(meta.PrimaryNameAttribute)) attributes.Add(GetFullAttributeName(meta.PrimaryNameAttribute, alias));
                attributes.AddRange(meta.Attributes.Select(q => GetFullAttributeName(q.LogicalName, alias)).OrderBy(q => q).Except(attributes));
                return attributes.ToArray();
            }
            else
            {
                return columnSet.Columns.Select(q => GetFullAttributeName(q, alias));
            }
        }

        private static string GetFullAttributeName(string attributeName, string alias)
        {
            return string.IsNullOrEmpty(alias) ? attributeName : string.Format("{0}.{1}", alias, attributeName);
        }
    }
}