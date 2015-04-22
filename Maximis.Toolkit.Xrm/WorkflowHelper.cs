using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Maximis.Toolkit.Xrm
{
    public static class WorkflowHelper
    {
        /// <summary>
        /// Cancels an instance of a Workflow
        /// </summary>
        public static void CancelWorkflow(IOrganizationService orgService, Guid asyncOperationId)
        {
            UpdateHelper.SetEntityState(orgService, new EntityReference("asyncoperation", asyncOperationId), 3, 32);
        }

        /// <summary>
        /// Retrieves a workflow's ID by name
        /// </summary>
        public static Guid GetWorkflowId(IOrganizationService orgService, string entityType, string workflowName)
        {
            QueryExpression query = new QueryExpression("workflow");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, workflowName);
            query.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
            query.Criteria.AddCondition("ondemand", ConditionOperator.Equal, true);
            query.Criteria.AddCondition("primaryentity", ConditionOperator.Equal, entityType);
            Entity workflow = QueryHelper.RetrieveSingleEntity(orgService, query);
            if (workflow == null) throw new Exception(string.Format("Workflow '{0}' does not exist", workflowName));
            return workflow.Id;
        }

        /// <summary>
        /// Postpones an instance of a Workflow
        /// </summary>
        public static void PostponeWorkflow(IOrganizationService orgService, Guid asyncOperationId, DateTime dateTime)
        {
            Entity asyncOp = new Entity("asyncoperation") { Id = asyncOperationId };
            asyncOp["postponeuntil"] = dateTime;
            orgService.Update(asyncOp);
        }

        /// <summary>
        /// Resumes a postponed instance of a Workflow
        /// </summary>
        public static void ResumePostponedWorkflow(IOrganizationService orgService, Guid asyncOperationId)
        {
            Entity asyncOp = new Entity("asyncoperation") { Id = asyncOperationId };
            asyncOp["postponeuntil"] = DateTime.Now;
            orgService.Update(asyncOp);
        }

        /// <summary>
        /// Retrieves a workflow by name and then runs it
        /// </summary>
        public static ExecuteWorkflowResponse RunWorkflow(IOrganizationService orgService,
            EntityReference entityReference, string workflowName)
        {
            Guid workflowId = GetWorkflowId(orgService, entityReference.LogicalName, workflowName);
            return RunWorkflow(orgService, entityReference, workflowId);
        }

        /// <summary>
        /// Triggers a Workflow against an Entity
        /// </summary>
        public static ExecuteWorkflowResponse RunWorkflow(IOrganizationService orgService,
            EntityReference entityReference, Guid workflowId)
        {
            ExecuteWorkflowRequest wfExecute = new ExecuteWorkflowRequest
            {
                WorkflowId = workflowId,
                EntityId = entityReference.Id
            };
            return (ExecuteWorkflowResponse)orgService.Execute(wfExecute);
        }
    }
}