using Maximis.Toolkit.Logging;
using Maximis.Toolkit.Xrm.EntitySerialisation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    internal class XmlImportManager : IDisposable
    {
        private OrganizationRequestCollection createUpdateRequests;
        private EntityDeserialiser deserialiser;
        private MetadataCache metaCache;
        private ImportOptions options;
        private IOrganizationService orgService;
        private OrganizationRequestCollection setStateRequests;

        public XmlImportManager(IOrganizationService orgService, MetadataCache metaCache, ImportOptions options)
        {
            this.orgService = orgService;
            this.metaCache = metaCache;
            this.options = options;
            this.deserialiser = new EntityDeserialiser(metaCache);

            this.createUpdateRequests = new OrganizationRequestCollection();
            this.setStateRequests = new OrganizationRequestCollection();
        }

        public void Dispose()
        {
            if (createUpdateRequests.Count > 0)
            {
                PerformImport();
            }
        }

        internal void AddForImport(string xml)
        {
            // Deserialise into Entity
            Entity newEntity = deserialiser.DeserialiseEntity(orgService, xml, options);

            // Drop out if Entity is empty
            if (newEntity.Id == Guid.Empty && !newEntity.Attributes.Any()) return;

            // Get Metadata for Entity
            EntityMetadata meta = metaCache.GetEntityMetadata(orgService, newEntity.LogicalName);

            // Check for Existing Record if required
            Entity existingEntity = null;
            if (options.CheckForExisting != CheckForExistingMode.DoNotCheck)
            {
                // Set Primary Id Attribute in case that is used for attribute matching
                if (newEntity.Id != Guid.Empty)
                {
                    newEntity[meta.PrimaryIdAttribute] = newEntity.Id;
                }

                QueryExpression existingQuery = new QueryExpression(newEntity.LogicalName) { ColumnSet = new ColumnSet(newEntity.Attributes.Keys.ToArray()) };

                switch (options.CheckForExisting)
                {
                    case CheckForExistingMode.Id:

                        existingQuery.Criteria.AddCondition(meta.PrimaryIdAttribute, ConditionOperator.Equal, newEntity.Id);
                        existingEntity = QueryHelper.RetrieveSingleEntity(orgService, existingQuery);
                        break;

                    case CheckForExistingMode.AllAttributes:
                    case CheckForExistingMode.AnyAttribute:

                        existingQuery.Criteria.FilterOperator = (options.CheckForExisting == CheckForExistingMode.AnyAttribute) ? LogicalOperator.Or : LogicalOperator.And;
                        IEnumerable<ExistingMatchAttribute> existingMatchAttributes;
                        int priority = 1;
                        while (true)
                        {
                            existingMatchAttributes = options.ExistingMatchAttributes.Where(q => q.Priority == priority);
                            if (!existingMatchAttributes.Any()) break;

                            existingQuery.Criteria.Conditions.Clear();
                            foreach (string attributeName in existingMatchAttributes.Select(q => q.AttributeName))
                            {
                                if (newEntity.Contains(attributeName))
                                {
                                    object val = newEntity[attributeName];
                                    if (val != null)
                                    {
                                        // TODO - Deal with Lookup, Money, etc where this simple operation won't work
                                        existingQuery.Criteria.AddCondition(attributeName, ConditionOperator.Equal, newEntity[attributeName]);
                                    }
                                }
                            }
                            if (existingQuery.Criteria.Conditions.Count > 0)
                            {
                                existingEntity = QueryHelper.RetrieveSingleEntity(orgService, existingQuery);
                                if (existingEntity != null) break;
                            }

                            ++priority;
                        }

                        break;
                }
            }

            // Make sure IDs are set correctly
            bool isNew = false;
            if (existingEntity == null)
            {
                if (newEntity.Id == Guid.Empty) newEntity.Id = Guid.NewGuid();
                isNew = true;
            }
            else
            {
                // Existing Entity might have a different ID if ID was not criteria used for existing match
                newEntity.Id = existingEntity.Id;
                if (newEntity.Contains(meta.PrimaryIdAttribute)) newEntity[meta.PrimaryIdAttribute] = newEntity.Id;
            }

            // If statecode and statuscode have been set, remove those values from the Entity and create a SetStateRequest instead
            if (newEntity.HasAttributeWithValue("statecode") && newEntity.HasAttributeWithValue("statuscode"))
            {
                setStateRequests.Add(new SetStateRequest
                {
                    EntityMoniker = newEntity.ToEntityReference(),
                    State = newEntity.GetAttributeValue<OptionSetValue>("statecode"),
                    Status = newEntity.GetAttributeValue<OptionSetValue>("statuscode")
                });
                newEntity.Attributes.Remove("statecode");
                newEntity.Attributes.Remove("statuscode");
            }

            // Create either CreateRequest or UpdateRequest
            if (isNew)
            {
                Trace.Write(string.Format("Creating new {0} record '{1}'...", newEntity.LogicalName, newEntity.Id));
                createUpdateRequests.Add(new CreateRequest { Target = newEntity });
            }
            else
            {
                // Filter down to only changed attributes
                newEntity = UpdateHelper.GetUpdateEntity(newEntity, existingEntity);

                if (newEntity.Attributes.Count > 0)
                {
                    Trace.Write(string.Format("Updating {0} attributes on {1} record '{2}'...", newEntity.Attributes.Count, newEntity.LogicalName, newEntity.Id));
                    createUpdateRequests.Add(new UpdateRequest { Target = newEntity });
                }
                else
                {
                    Trace.Write(string.Format("No change to existing {0} record '{1}'...", newEntity.LogicalName, existingEntity.Id));
                }
            }
            Trace.WriteLine("Done.");

            // Apply updates to CRM
            if (createUpdateRequests.Count >= options.BatchSize)
            {
                PerformImport();
            }
        }

        private void PerformImport()
        {
            using (TraceProgressReporter progress = new TraceProgressReporter("Writing to CRM"))
            {
                BulkOperationHelper.PerformExecuteMultiple(orgService, createUpdateRequests, createUpdateRequests.Count, progress, options.ContinueOnError);
                BulkOperationHelper.PerformExecuteMultiple(orgService, setStateRequests, setStateRequests.Count, progress, options.ContinueOnError);
            }
            createUpdateRequests.Clear();
            setStateRequests.Clear();
        }
    }
}