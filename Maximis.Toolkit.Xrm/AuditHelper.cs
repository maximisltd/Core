using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm
{
    public static class AuditHelper
    {
        public static List<Entity> GetAuditRecords(IOrganizationService orgService, EntityReference entityRef)
        {
            QueryExpression query = new QueryExpression("audit") { ColumnSet = new ColumnSet(true) };
            query.Criteria.AddCondition("objectid", ConditionOperator.Equal, entityRef.Id);
            query.Criteria.AddCondition("changedata", ConditionOperator.NotNull);
            return QueryHelper.RetrieveAllEntities(orgService, query);
        }
    }
}