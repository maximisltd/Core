using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;

namespace Maximis.Toolkit.Xrm
{
    public abstract class BaseCustomWorkflowActivity : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Get the Workflow Context
            IWorkflowContext workflowContext = executionContext.GetExtension<IWorkflowContext>();

            // Get the Organization Service
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.UserId);

            // Get the Tracing Service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            // Call ExecuteWorkflowStep method
            ExecuteWorkflowStep(workflowContext, service, tracingService, executionContext);
        }

        protected abstract void ExecuteWorkflowStep(IWorkflowContext workflowContext, IOrganizationService orgService,
            ITracingService tracingService, CodeActivityContext executionContext);
    }
}