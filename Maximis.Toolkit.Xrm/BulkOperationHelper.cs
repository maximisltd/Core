using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Maximis.Toolkit.Xrm
{
    public enum DeleteMode { Execute, ExecuteMultiple, BulkDeleteJob }

    public static class BulkOperationHelper
    {
        public static void DeleteEntities(CrmContext context, QueryExpression query, DeleteMode deleteMode = DeleteMode.ExecuteMultiple, int execMultipleBatchSize = 200)
        {
            if (deleteMode == DeleteMode.BulkDeleteJob)
            {
                context.OrganizationService.Execute(new BulkDeleteRequest
                {
                    JobName = string.Format("Bulk Delete: {0}", query.EntityName),
                    RecurrencePattern = string.Empty,
                    SendEmailNotification = false,
                    ToRecipients = new Guid[] { },
                    CCRecipients = new Guid[] { },
                    QuerySet = new[] { query }
                });
            }
            else
            {
                List<Entity> toDeleteList = QueryHelper.RetrieveAllEntities(context.OrganizationService, query);
                if (!toDeleteList.Any()) return;

                switch (deleteMode)
                {
                    case DeleteMode.ExecuteMultiple:
                        OrganizationRequestCollection deleteRequests = new OrganizationRequestCollection();
                        deleteRequests.AddRange(toDeleteList.Select(q => new DeleteRequest { Target = q.ToEntityReference() }));
                        PerformExecuteMultiple(context, deleteRequests, execMultipleBatchSize, true);
                        break;

                    case DeleteMode.Execute:
                        foreach (Entity toDelete in toDeleteList)
                        {
                            context.OrganizationService.Delete(toDelete.LogicalName, toDelete.Id);
                        }
                        break;
                }
            }
        }

        public static void DeleteEntities(CrmContext context, string entityType, DeleteMode deleteMode = DeleteMode.ExecuteMultiple, int execMultipleBatchSize = 200)
        {
            DeleteEntities(context, new QueryExpression(entityType), deleteMode, execMultipleBatchSize);
        }

        public static bool PerformExecuteMultiple(CrmContext context, IEnumerable<OrganizationRequest> requests, int batchSize = 0, bool continueOnError = false)
        {
            // Drop out if Request Collection is empty
            if (!requests.Any()) return true;

            // Split Requests into batches if a Batch Size was supplied

            bool result = true;
            foreach (IEnumerable<OrganizationRequest> requestBatch in requests.InSetsOf(batchSize > 0 ? batchSize : requests.Count()))
            {
                // Create Request Collection
                OrganizationRequestCollection requestBatchCollection = new OrganizationRequestCollection();
                requestBatchCollection.AddRange(requestBatch);

                // Perform Execute Multiple
                ExecuteMultipleResponse rsp = (ExecuteMultipleResponse)context.OrganizationService.Execute(new ExecuteMultipleRequest
                {
                    Requests = requestBatchCollection,
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = continueOnError,
                        ReturnResponses = true
                    }
                });

                // Handle Faults
                if (rsp.IsFaulted)
                {
                    HandleFaults(rsp, requestBatchCollection, context, continueOnError);
                    result = false;
                }
            }
            return result;
        }

        private static void HandleFaults(ExecuteMultipleResponse rsp, OrganizationRequestCollection requestBatchCollection, CrmContext context, bool continueOnError)
        {
            StringBuilder error = new StringBuilder();

            // Loop through Faults (there will only be one if continueOnError is false)
            foreach (ExecuteMultipleResponseItem failure in rsp.Responses.Where(q => q.Fault != null))
            {
                // Find the Request
                OrganizationRequest request = requestBatchCollection[failure.RequestIndex];

                // Build an error message string
                error.AppendLine();
                error.AppendLine("==========");
                error.Append("ERROR - ");
                error.AppendLine(request.GetType().Name);

                // Add info about the Fault to the error message
                OrganizationServiceFault fault = failure.Fault;
                for (int i = 0; i < 10; i++)
                {
                    error.Append(" -- ");
                    error.AppendLine(fault.Message);
                    if (fault.InnerFault == null) break;
                    fault = fault.InnerFault;
                }

                // Look for an Entity
                foreach (KeyValuePair<string, object> parameter in request.Parameters)
                {
                    // Add Entity Type and ID to error message
                    Entity entity = parameter.Value as Entity;
                    if (entity != null)
                    {
                        error.Append("ENTITY: ");
                        error.Append(entity.LogicalName);
                        error.Append(" ");
                        error.AppendLine(entity.Id.ToString("D"));
                    }
                }

                error.AppendLine();
                error.AppendLine("==========");
                error.AppendLine();

                // Output error message to Tracing service
                context.TracingService.Trace(error.ToString());
            }

            if (!continueOnError) throw new InvalidPluginExecutionException(error.ToString());
        }
    }
}