using Maximis.Toolkit.Caching;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Maximis.Toolkit.Xrm
{
    public class WorkflowStepContext : CrmContext
    {
        public WorkflowStepContext(CodeActivityContext activityContext, Type type)
        {
            this.ActivityContext = activityContext;
            this.WorkflowContext = activityContext.GetExtension<IWorkflowContext>();
            this.TracingService = new TracingService(activityContext.GetExtension<ITracingService>());
            this.CacheManager = new CacheManager(this.WorkflowContext.OrganizationName);
            this.OrganizationService = activityContext.GetExtension<IOrganizationServiceFactory>().CreateOrganizationService(this.WorkflowContext.UserId);
            this.ExecutionContext = this.WorkflowContext;
            this.TracingService.Trace("Custom Workflow Step Instantiated: {0}", type.Name);
            this.TracingService.Trace("Primary Record: {0} {1}", this.WorkflowContext.PrimaryEntityName, this.WorkflowContext.PrimaryEntityId);
            this.TracingService.Trace("Depth: {0}", this.WorkflowContext.Depth);
        }

        public CodeActivityContext ActivityContext { get; set; }

        public IWorkflowContext WorkflowContext { get; set; }
    }
}