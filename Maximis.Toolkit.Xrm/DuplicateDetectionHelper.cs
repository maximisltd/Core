using Maximis.Toolkit.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel;

namespace Maximis.Toolkit.Xrm
{
    public static class DuplicateDetectionHelper
    {
        public static int CountDuplicates(IOrganizationService orgService, Entity entity)
        {
            return RetrieveDuplicates(orgService, entity).Entities.Count;
        }

        public static void PublishAllRules(IOrganizationService orgService)
        {
            using (TraceProgressReporter progress = new TraceProgressReporter("Publishing Duplicate Detection Rules", 1))
            {
                QueryExpression query = new QueryExpression { EntityName = "duplicaterule" };
                query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 0);
                foreach (Entity dupeRulEntity in orgService.RetrieveMultiple(query).Entities)
                {
                    try
                    {
                        orgService.Execute(new PublishDuplicateRuleRequest() { DuplicateRuleId = dupeRulEntity.Id });
                        progress.IterationComplete();
                    }
                    catch (FaultException ex)
                    {
                        // Suppress message
                        if (ex.Message.StartsWith("Duplicate detection is not enabled"))
                        {
                            progress.OperationFailed(ex.Message);
                            return;
                        }
                        else
                        {
                            progress.IterationFailed(ex.Message);
                        }
                    }
                }
            }
        }

        public static EntityCollection RetrieveDuplicates(IOrganizationService orgService, Entity entity)
        {
            RetrieveDuplicatesResponse dupRsp = (RetrieveDuplicatesResponse)orgService.Execute(
                new RetrieveDuplicatesRequest
                {
                    BusinessEntity = entity,
                    MatchingEntityName = entity.LogicalName,
                    PagingInfo = new PagingInfo { PageNumber = 1, Count = 1 }
                });
            return dupRsp.DuplicateCollection;
        }
    }
}