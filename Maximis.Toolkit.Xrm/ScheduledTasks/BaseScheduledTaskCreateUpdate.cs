using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Maximis.Toolkit.Xrm.ScheduledTasks
{
    public abstract class BaseScheduledTaskCreateUpdate : BasePlugin
    {
        private static readonly string jobName = "Delete Scheduled Task records (causes workflow to trigger)";

        public BaseScheduledTaskCreateUpdate(string nameField, string workflowNameField, string dateField)
        {
            NameField = nameField;
            WorkflowNameField = workflowNameField;
            DateField = dateField;
        }

        protected string DateField { get; set; }

        protected string NameField { get; set; }

        protected string WorkflowNameField { get; set; }

        /// <summary>
        /// Ensures that Bulk Delete jobs are set up and the record has a meaningful name.
        /// </summary>
        protected override void ExecutePlugin(IServiceProvider serviceProvider, IPluginExecutionContext context,
            ITracingService tracingService)
        {
            // Get Scheduled Task Entity
            Entity taskRecord = GetParameter<Entity>(context);
            if (context.MessageName == "Update")
                taskRecord = taskRecord.Merge(GetEntityImage(context, "ScheduledTask", EntityImageType.PreOperation));

            // Set name
            taskRecord[NameField] = string.Format("Run workflow '{0}' on or after {1:dd MMM yyyy HH:mm}",
                taskRecord[WorkflowNameField], taskRecord[DateField]);

            // Retrieve bulk delete jobs
            IOrganizationService orgService = GetOrganizationService(serviceProvider, context);
            QueryExpression existingQuery = new QueryExpression("asyncoperation") { ColumnSet = new ColumnSet(true) };
            existingQuery.Criteria.AddCondition("operationtype", ConditionOperator.Equal, 13);
            existingQuery.Criteria.AddCondition("name", ConditionOperator.EndsWith, jobName);
            EntityCollection existingJobs = orgService.RetrieveMultiple(existingQuery);

            // If jobs exist, we know we don't need to run code below to set them up
            if (existingJobs.Entities.Count > 0)
            {
                // Delete successful jobs older than a week
                DateTime oneWeekAgo = DateTime.Now.AddDays(-7);
                IEnumerable<Entity> toDelete =
                    existingJobs.Entities.Where(q => q.GetLocalDateTimeValue("startedon") < oneWeekAgo
                                                     && q.GetAttributeValue<OptionSetValue>("statuscode").Value == 30);
                foreach (Entity existingJob in toDelete)
                {
                    orgService.Delete(existingJob.LogicalName, existingJob.Id);
                }

                // Return so no further code is run
                return;
            }

            // Query to select Scheduled Task entities with "Run on or after" date in the past
            QueryExpression qe = new QueryExpression(context.PrimaryEntityName);
            qe.Criteria.AddCondition(new ConditionExpression(DateField, ConditionOperator.LastXYears, 10));
            qe.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            // Set up 24 jobs, one per hour.
            DateTime startDateTime = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day,
                DateTime.Today.Hour, 0, 0).AddHours(1);

            for (int i = 1; i <= 24; i++)
            {
                BulkDeleteRequest bulkDeleteRequest = new BulkDeleteRequest
                {
                    JobName = string.Format("[{0:HH:mm}] {1}", startDateTime, jobName),
                    RecurrencePattern = "FREQ=DAILY;INTERVAL=1;",
                    StartDateTime = startDateTime,
                    SendEmailNotification = false,
                    ToRecipients = new Guid[] { },
                    CCRecipients = new Guid[] { },
                    QuerySet = new[] { qe }
                };

                orgService.Execute(bulkDeleteRequest);

                startDateTime = startDateTime.AddHours(1);
            }
        }
    }
}