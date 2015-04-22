using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Maximis.Toolkit.Xrm.ScheduledTasks
{
    /// <summary>
    /// Triggers the workflow defined in the Scheduled Task entity when it is deleted by a Bulk
    /// Delete job
    /// </summary>
    public abstract class BaseScheduledTaskDelete : BasePlugin
    {
        public BaseScheduledTaskDelete(string workflowNameField, string dateField, string frequencyField)
        {
            WorkflowNameField = workflowNameField;
            DateField = dateField;
            FrequencyField = frequencyField;
        }

        public string DateField { get; set; }

        public string FrequencyField { get; set; }

        protected string WorkflowNameField { get; set; }

        protected override void ExecutePlugin(IServiceProvider serviceProvider, IPluginExecutionContext context,
            ITracingService tracingService)
        {
            // Get the Scheduled Task entity being deleted
            EntityReference targetRef = GetParameter<EntityReference>(context);
            IOrganizationService service = GetOrganizationService(serviceProvider, context);
            Entity taskBeingDeleted = service.Retrieve(targetRef.LogicalName, targetRef.Id,
                new ColumnSet(WorkflowNameField, DateField, FrequencyField, "statecode"));

            // If Entity is inactive, do not trigger workflow and do not re-create a new Entity
            if (taskBeingDeleted.GetAttributeValue<OptionSetValue>("statecode").Value != 0) return;

            // Get the name of the Workflow which is to be fired
            string workflowName = taskBeingDeleted.GetAttributeValue<string>(WorkflowNameField);

            // Get the run time of the Custom Task which is to be fired
            DateTime runTime = taskBeingDeleted.GetLocalDateTimeValue(DateField);

            // Get the frequency at which jobs are run
            decimal frequency = taskBeingDeleted.GetAttributeValue<decimal>(FrequencyField);
            double frequencyDbl = (double)frequency;

            // Create a new Scheduled Task entity which will be deleted at the right time
            DateTime nextJobDateTime = runTime;
            while (nextJobDateTime < DateTime.Now)
            {
                nextJobDateTime = nextJobDateTime.AddHours(frequencyDbl);
            }
            Entity newTask = new Entity(targetRef.LogicalName);
            newTask.Attributes[WorkflowNameField] = workflowName;
            newTask.Attributes[DateField] = nextJobDateTime;
            newTask.Attributes[FrequencyField] = frequency;
            newTask.Id = service.Create(newTask);
            Trace(tracingService, "New Scheduled Task record created to be picked up by next delete job");

            // Trigger Workflow - has to be done against newly created Scheduled Task entity
            if (!string.IsNullOrEmpty(workflowName))
            {
                Trace(tracingService, "Scheduled Task record deleted - attempting to trigger Workflow '{0}'",
                    workflowName);
                ExecuteWorkflowResponse workflowResponse = WorkflowHelper.RunWorkflow(service,
                    newTask.ToEntityReference(), workflowName);
                if (workflowResponse != null)
                {
                    Trace(tracingService, "Workflow '{0}' triggered successfully. System Job ID: '{1:B}'", workflowName,
                        workflowResponse.Id);
                }
                else
                {
                    Trace(tracingService,
                        "Could not find Workflow '{0}'. Check the spelling, that the Workflow exists and is activated. Check it can be called on demand and that its Primary Entity is 'Scheduled Task'.",
                        workflowName);
                }
            }
            else
            {
                Trace(tracingService,
                    "Scheduled Task record deleted - no Workflow name defined, so no action will be taken.");
            }
        }
    }
}