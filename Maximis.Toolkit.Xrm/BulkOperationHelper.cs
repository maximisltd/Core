using Maximis.Toolkit.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Maximis.Toolkit.Xrm
{
    public enum DeleteMode { Execute, ExecuteMultiple, BulkDeleteJob }

    public static class BulkOperationHelper
    {
        public static void DeleteEntities(IOrganizationService orgService, string entityType, DeleteMode deleteMode = DeleteMode.ExecuteMultiple)
        {
            QueryExpression query = new QueryExpression(entityType);
            DeleteEntities(orgService, query, deleteMode);
        }

        public static void DeleteEntities(IOrganizationService orgService, QueryExpression query, DeleteMode deleteMode = DeleteMode.ExecuteMultiple, int execMultipleBatchSize = 200)
        {
            if (deleteMode == DeleteMode.BulkDeleteJob)
            {
                using (TraceProgressReporter progress = new TraceProgressReporter(string.Format("Creating Bulk Delete job for '{0}' records", query.EntityName)))
                {
                    orgService.Execute(new BulkDeleteRequest
                    {
                        JobName = string.Format("Bulk Delete: {0}", query.EntityName),
                        RecurrencePattern = string.Empty,
                        SendEmailNotification = false,
                        ToRecipients = new Guid[] { },
                        CCRecipients = new Guid[] { },
                        QuerySet = new[] { query }
                    });
                }
            }
            else
            {
                List<Entity> toDeleteList = null;
                using (TraceProgressReporter progress = new TraceProgressReporter(string.Format("Retrieving '{0}' records for deletion", query.EntityName)))
                {
                    toDeleteList = QueryHelper.RetrieveAllEntities(orgService, query);
                }
                using (TraceProgressReporter progress = new TraceProgressReporter(string.Format("Deleting {0} '{1}' records", toDeleteList.Count, query.EntityName), 1))
                {
                    switch (deleteMode)
                    {
                        case DeleteMode.ExecuteMultiple:
                            OrganizationRequestCollection deleteRequests = new OrganizationRequestCollection();
                            deleteRequests.AddRange(toDeleteList.Select(q => new DeleteRequest { Target = q.ToEntityReference() }));
                            PerformExecuteMultiple(orgService, deleteRequests, execMultipleBatchSize, progress, true);
                            break;

                        case DeleteMode.Execute:
                            foreach (Entity toDelete in toDeleteList)
                            {
                                orgService.Delete(toDelete.LogicalName, toDelete.Id);
                                progress.IterationComplete();
                            }
                            break;
                    }
                }
            }
        }

        public static void PerformExecuteMultiple(IOrganizationService orgService, IEnumerable<OrganizationRequest> requests, int batchSize = 0, TraceProgressReporter progress = null, bool continueOnError = false)
        {
            if (!requests.Any()) return;

            // Get OrganizationServiceProxy object (will be null if orgService is a different type of object)
            OrganizationServiceProxy orgServiceProxy = orgService as OrganizationServiceProxy;

            MetadataCache metaCache = new MetadataCache();

            foreach (IEnumerable<OrganizationRequest> requestBatch in requests.InSetsOf(batchSize > 0 ? batchSize : requests.Count()))
            {
                OrganizationRequestCollection requestBatchCollection = new OrganizationRequestCollection();
                requestBatchCollection.AddRange(requestBatch);

                ExecuteMultipleResponse rsp = (ExecuteMultipleResponse)orgService.Execute(new ExecuteMultipleRequest
                {
                    Requests = requestBatchCollection,
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = continueOnError,
                        ReturnResponses = true
                    }
                });

                if (rsp.IsFaulted)
                {
                    foreach (ExecuteMultipleResponseItem failure in rsp.Responses.Where(q => q.Fault != null))
                    {
                        OrganizationRequest request = requestBatchCollection[failure.RequestIndex];

                        StringBuilder error = new StringBuilder();
                        error.AppendLine();
                        error.AppendLine("==========");
                        error.Append("ERROR - ");
                        error.AppendLine(request.GetType().Name);

                        OrganizationServiceFault fault = failure.Fault;
                        for (int i = 0; i < 10; i++)
                        {
                            error.Append(" -- ");
                            error.AppendLine(fault.Message);
                            if (fault.InnerFault == null) break;
                            fault = fault.InnerFault;
                        }

                        CreateRequest createRequest = request as CreateRequest;
                        if (createRequest != null)
                        {
                            OutputErrorEntity(orgService, error, createRequest.Target, metaCache);
                        }

                        UpdateRequest updateRequest = request as UpdateRequest;
                        if (updateRequest != null)
                        {
                            OutputErrorEntity(orgService, error, updateRequest.Target, metaCache);
                        }

                        error.AppendLine();
                        error.AppendLine("==========");
                        error.AppendLine();

                        progress.WriteInfo(error.ToString());

                        if (!continueOnError) throw new FaultException<OrganizationServiceFault>(failure.Fault);
                    }
                }

                if (progress != null) progress.IterationComplete();

                // Re-authenticate if necessary
                if (orgServiceProxy != null)
                {
                    ServiceHelper.RenewTokenIfRequired(orgServiceProxy);
                }
            }
        }

        private static void OutputErrorEntity(IOrganizationService orgService, StringBuilder error, Entity entity, MetadataCache metaCache)
        {
            error.AppendLine();
            error.AppendLine("ENTITY VALUES");
            error.AppendLine(string.Format("Entity Id: '{0:D}'", entity.Id));
            foreach (string attribute in entity.Attributes.Keys)
            {
                error.AppendLine(string.Format("{0}: '{1}'", attribute, MetadataHelper.GetAttributeValueAsDisplayString(orgService, metaCache, entity, attribute, new DisplayStringOptions())));
            }
        }
    }
}