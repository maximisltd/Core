using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Maximis.Toolkit.Xrm
{
    public static class UpdateHelper
    {
        /// <summary>
        /// Calls the Create or Update method of IOrganizationService as required
        /// </summary>
        public static Entity CreateOrUpdate(IOrganizationService orgService, Entity entity, bool forceCreate = false, bool checkForDuplicatesOnCreate = false)
        {
            // If State/Status are included in the entity, extract them and deal with them separately
            bool updateStatus = false;
            int newStateCode = 0;
            int newStatusCode = 0;
            if (entity.HasAttributeWithValue("statecode") && entity.HasAttributeWithValue("statuscode"))
            {
                updateStatus = true;
                newStateCode = entity.GetAttributeValue<OptionSetValue>("statecode").Value;
                newStatusCode = entity.GetAttributeValue<OptionSetValue>("statuscode").Value;
                entity.Attributes.Remove("statecode");
                entity.Attributes.Remove("statuscode");
            }

            // Create or Update the Entity
            if (forceCreate || entity.Id == Guid.Empty)
            {
                if (checkForDuplicatesOnCreate && DuplicateDetectionHelper.CountDuplicates(orgService, entity) > 0)
                {
                    throw new DuplicateRecordException("Record not created - CRM Duplicate Detection rules detected a duplicate record.");
                }
                entity.Id = orgService.Create(entity);
            }
            else
            {
                orgService.Update(entity);
            }

            // Update the State/Status if required
            if (updateStatus)
            {
                SetEntityState(orgService, entity.ToEntityReference(), newStateCode, newStatusCode);
            }

            return entity;
        }

        /// <summary>
        /// Takes two versions of the same Entity and returns an entity which contains only the differences, used to apply an Update that won't create unnecessary audit records.
        /// </summary>
        public static Entity GetUpdateEntity(Entity update, Entity current)
        {
            if (update.Id != current.Id) throw new Exception("Entity Ids do not match!");

            if (update.Attributes.Count > 0)
            {
                string[] keysArray = update.Attributes.Keys.ToArray();
                foreach (string key in keysArray)
                {
                    if (current.HasAttributeWithValue(key))
                    {
                        // Original Entity has a value for this attribute

                        // Read original and new values
                        object oldVal = current[key];
                        object newVal = update[key];

                        if (oldVal is DateTime)
                        {
                            // If DateTime values are the same (accounting for local time), do not update
                            DateTime oldDate = ((DateTime)oldVal).ToUniversalTime();
                            if (newVal != null)
                            {
                                DateTime newDate = ((DateTime)newVal).ToUniversalTime();
                                if (newDate == oldDate)
                                {
                                    update.Attributes.Remove(key);
                                }
                            }
                        }
                        else if (oldVal is EntityReference)
                        {
                            // If new EntityReference has same ID as old, do not update
                            EntityReference oldRef = oldVal as EntityReference;
                            EntityReference newRef = newVal as EntityReference;
                            if (newRef != null && newRef.Id == oldRef.Id)
                            {
                                update.Attributes.Remove(key);
                            }
                        }
                        else if (oldVal is Money)
                        {
                            if (newVal != null)
                            {
                                Money oldMoney = (Money)oldVal;
                                Money newMoney = (Money)newVal;
                                if (newMoney.Value == oldMoney.Value)
                                {
                                    update.Attributes.Remove(key);
                                }
                            }
                        }
                        else if (oldVal is string)
                        {
                            string oldStr = (string)oldVal;
                            string newStr = (string)newVal;

                            if (string.IsNullOrEmpty(oldStr) && string.IsNullOrEmpty(newStr))
                            {
                                // Explicity set to null even if either is String.Empty
                                update.Attributes[key] = null;
                            }
                            else if (oldStr == newStr)
                            {
                                update.Attributes.Remove(key);
                            }
                        }
                        else if (oldVal.Equals(newVal))
                        {
                            // If other type of values are the same, do not update this value
                            update.Attributes.Remove(key);
                        }
                    }
                    else
                    {
                        // Does not exist in original, so do not update if updated value is NULL
                        if (update[key] == null || string.IsNullOrEmpty(update[key].ToString()))
                        {
                            update.Attributes.Remove(key);
                        }
                    }
                }
            }

            return update;
        }

        /// <summary>
        /// Implements a relationship between one entity and one or several others. This can be a
        /// one-to-many or many-to-many relationship
        /// </summary>
        public static void RelateEntities(IOrganizationService orgService, string relationshipName,
            EntityReference relateFrom, params EntityReference[] relateTo)
        {
            try
            {
                orgService.Associate(relateFrom.LogicalName, relateFrom.Id, new Relationship(relationshipName),
                    new EntityReferenceCollection(relateTo));
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("duplicate")) throw ex;
            }
        }

        /// <summary>
        /// Finds a permitted Relationship between two Entites and implements it. This is inefficient: do not use in Production code. Useful for setup scripts, data migrations, etc.
        /// </summary>
        public static void RelateEntitiesLazy(IOrganizationService orgService, EntityReference relateFrom, EntityReference relateTo)
        {
            RelationshipMetadataBase relationship = MetadataHelper.GetAllRelationships(orgService, relateFrom.LogicalName, relateTo.LogicalName).FirstOrDefault();
            if (relationship != null)
            {
                RelateEntities(orgService, relationship.SchemaName, relateFrom, relateTo);
            }
        }

        /// <summary>
        /// Changes the State and Status of an Entity to the supplied values
        /// </summary>
        public static void SetEntityState(IOrganizationService orgService, EntityReference entityReference,
            int stateCode, int statusCode)
        {
            orgService.Execute(new SetStateRequest
            {
                EntityMoniker = entityReference,
                State = new OptionSetValue(stateCode),
                Status = new OptionSetValue(statusCode)
            });
        }

        /// <summary>
        /// Only applies an update to fields which are different from the current values
        /// </summary>
        public static int SmartUpdate(IOrganizationService orgService, Entity update, Entity current = null)
        {
            // Retrieve current version of record if not supplied
            if (current == null && update.Id != Guid.Empty)
            {
                ColumnSet colSet = new ColumnSet(update.Attributes.Select(q => q.Key).ToArray());
                current = QueryHelper.RetrieveSingleEntity(orgService, update.LogicalName, update.Id, colSet);
            }

            if (current == null)
            {
                // Create new record
                if (update.Attributes.Count > 0)
                {
                    UpdateHelper.CreateOrUpdate(orgService, update, true);
                }
            }
            else
            {
                // Update record
                if (update.Attributes.Count > 0)
                {
                    update = GetUpdateEntity(update, current);
                    UpdateHelper.CreateOrUpdate(orgService, update);
                }
            }
            return update.Attributes.Count;
        }
    }
}