using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Maximis.Toolkit.Xrm
{
    public static class QueryHelper
    {
        /// <summary>
        /// Returns true if two Entity References are related via the given relationship
        /// </summary>
        public static bool AreRelated(IOrganizationService orgService, RelationshipMetadataBase relationshipMetadata, EntityReference entityRefA, EntityReference entityRefB)
        {
            EntityCollection relatedEntities = RetrieveRelatedEntities(orgService, entityRefA, relationshipMetadata);
            return relatedEntities.Entities.Any(q => q.LogicalName == entityRefB.LogicalName && q.Id == entityRefB.Id);
        }

        /// <summary>
        /// Counts the number of Entities returned by a query. Set usePaging if the result might be
        /// greater than 5,000
        /// </summary>
        public static int CountEntities(IOrganizationService orgService, QueryExpression query, bool usePaging = false)
        {
            // Make sure no attributes are returned
            query.ColumnSet = null;

            if (usePaging)
            {
                // Use RetrieveEntitiesWithPaging method to count
                int count = 0;
                EntityCollection results = null;

                while (RetrieveEntitiesWithPaging(orgService, query, ref results, 5000))
                {
                    count += results.Entities.Count;
                }
                return count;
            }
            else
            {
                // Use RetrieveEntities method to count
                return orgService.RetrieveMultiple(query).Entities.Count;
            }
        }

        public static Guid GetCurrentUserId(IOrganizationService orgService)
        {
            WhoAmIResponse response = orgService.Execute(new WhoAmIRequest()) as WhoAmIResponse;
            return response.UserId;
        }

        /// <summary>
        /// Returns a list of all Entities returned by a query, with no upper limit (uses paging)
        /// </summary>
        public static List<Entity> RetrieveAllEntities(IOrganizationService orgService, QueryExpression query)
        {
            List<Entity> result = new List<Entity>();
            EntityCollection ec = null;
            while (RetrieveEntitiesWithPaging(orgService, query, ref ec))
            {
                result.AddRange(ec.Entities);
            }
            return result;
        }

        /// <summary>
        /// Returns a collection of Entity objects. Handles paging when retrieving a large number of records.
        /// </summary>
        public static bool RetrieveEntitiesWithPaging(IOrganizationService orgService, QueryExpression query,
            ref EntityCollection results, int recordsPerPage = 500)
        {
            // Drop out if all records have been returned
            if (results == null)
            {
                // First time in - set up paging
                query.PageInfo = new PagingInfo { Count = recordsPerPage, PageNumber = 1 };
            }
            else
            {
                // Return false if no more data, otherwise update paging
                if (!results.MoreRecords) return false;
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = results.PagingCookie;
            }

            results = orgService.RetrieveMultiple(query);
            return (results.Entities != null && results.Entities.Count > 0);
        }

        /// <summary>
        /// Looks for an entity that matches the Query. Creates a new one if not found.
        /// </summary>
        public static Entity RetrieveOrNewEntity(IOrganizationService orgService, QueryExpression query)
        {
            Entity e = RetrieveSingleEntity(orgService, query);
            if (e == null)
            {
                e = new Entity(query.EntityName);
            }
            return e;
        }

        /// <summary>
        /// Returns all records of type "referencingEntityType" which have a lookup field
        /// ("lookupFieldName") set to referencedEntityId.
        /// </summary>
        public static EntityCollection RetrieveRelatedEntities(IOrganizationService orgService,
            string referencingEntityType, string lookupFieldName, Guid referencedEntityId, ColumnSet cols = null)
        {
            QueryExpression query = new QueryExpression(referencingEntityType) { ColumnSet = cols };
            query.Criteria.AddCondition(lookupFieldName, ConditionOperator.Equal, referencedEntityId);
            return orgService.RetrieveMultiple(query);
        }

        /// <summary>
        /// Returns all records related to "entityReference" by the relationship defined in
        /// "relationshipMetadata". Use MetadataHelper.GetRelationshipMetadata to get "relationshipMetadata".
        /// </summary>
        public static EntityCollection RetrieveRelatedEntities(IOrganizationService orgService,
            EntityReference entityReference, RelationshipMetadataBase relationshipMetadata, ColumnSet cols = null, params OrderExpression[] orders)
        {
            // Determine type (1:N or N:N)
            switch (relationshipMetadata.RelationshipType)
            {
                // One to Many
                case RelationshipType.OneToManyRelationship:

                    // The REFERENCED entity is the "ONE". The REFERENCING entity is the "MANY". The
                    // REFERENCING entity has a lookup field REFERENCING the REFERENCED entity.

                    // Cast Metadata
                    OneToManyRelationshipMetadata relOneMany = relationshipMetadata as OneToManyRelationshipMetadata;

                    // Operation differs depending on whether we have been passed the REFERENCED
                    // (One) or REFERENCING (Many) entity
                    if (entityReference.LogicalName == relOneMany.ReferencedEntity)
                    {
                        // If "entityReference" is the Referenced Entity, return all Referencing
                        // Entities which have a Lookup set to "entityReference"
                        return RetrieveRelatedEntities(orgService, relOneMany.ReferencingEntity, relOneMany.ReferencingAttribute,
                            entityReference.Id, cols);
                    }
                    else
                    {
                        // If "entityReference" is the Referencing Entity, return all Referenced entities
                        RetrieveRequest req = new RetrieveRequest
                        {
                            Target = entityReference,
                            ColumnSet = new ColumnSet(),
                            RelatedEntitiesQuery = new RelationshipQueryCollection()
                        };

                        QueryExpression relatedQuery = new QueryExpression(relOneMany.ReferencedEntity) { ColumnSet = cols };
                        if (orders != null && orders.Length > 0) relatedQuery.Orders.AddRange(orders);
                        req.RelatedEntitiesQuery.Add(new Relationship(relOneMany.SchemaName), relatedQuery);
                        RelatedEntityCollection result = ((RetrieveResponse)orgService.Execute(req)).Entity.RelatedEntities;
                        return (result == null || result.Count == 0) ? null : result.First().Value;
                    }

                // Many to Many
                case RelationshipType.ManyToManyRelationship:

                    // Many to Many relationships are really 2x One to Many with an "Intersect
                    // Entity" in between: A -< X >- B If "entityReference" is an A, we want to
                    // return all B's related to all X's related to A. If "entityReference" is a B,
                    // we want to return all A's related to all X's related to B.

                    // Cast Metadata
                    ManyToManyRelationshipMetadata relManyMany = relationshipMetadata as ManyToManyRelationshipMetadata;

                    // Create Query
                    QueryExpression query = new QueryExpression() { ColumnSet = cols };
                    if (orders != null && orders.Length > 0) query.Orders.AddRange(orders);

                    // Configure Query to return the "other type" of Entity from that of "entityReference"
                    if (entityReference.LogicalName == relManyMany.Entity1LogicalName)
                    {
                        query.EntityName = relManyMany.Entity2LogicalName;
                        LinkEntity linkEntity = new LinkEntity
                        {
                            LinkFromEntityName = relManyMany.Entity2LogicalName,
                            LinkFromAttributeName = relManyMany.Entity2IntersectAttribute,
                            LinkToEntityName = relManyMany.IntersectEntityName,
                            LinkToAttributeName = relManyMany.Entity2IntersectAttribute,
                        };
                        linkEntity.LinkCriteria.AddCondition(relManyMany.Entity1IntersectAttribute, ConditionOperator.Equal, entityReference.Id);
                        query.LinkEntities.Add(linkEntity);
                    }
                    else
                    {
                        query.EntityName = relManyMany.Entity1LogicalName;
                        LinkEntity linkEntity = new LinkEntity
                        {
                            LinkFromEntityName = relManyMany.Entity1LogicalName,
                            LinkFromAttributeName = relManyMany.Entity1IntersectAttribute,
                            LinkToEntityName = relManyMany.IntersectEntityName,
                            LinkToAttributeName = relManyMany.Entity1IntersectAttribute,
                        };
                        linkEntity.LinkCriteria.AddCondition(relManyMany.Entity2IntersectAttribute, ConditionOperator.Equal, entityReference.Id);
                        query.LinkEntities.Add(linkEntity);
                    }

                    // Return results of query
                    return orgService.RetrieveMultiple(query);
            }
            return null;
        }

        /// <summary>
        /// Returns a single record with a known Id
        /// </summary>
        public static Entity RetrieveSingleEntity(IOrganizationService orgService, string entityType, Guid Id,
            ColumnSet columnSet = null)
        {
            if (columnSet == null) columnSet = new ColumnSet();
            try { return orgService.Retrieve(entityType, Id, columnSet); }
            catch (Exception) { }
            return null;
        }

        /// <summary>
        /// Returns a single record by name
        /// </summary>
        public static Entity RetrieveSingleEntity(IOrganizationService orgService, string entityType, string primaryName,
            string primaryNameAttribute = null, ColumnSet columnSet = null)
        {
            if (string.IsNullOrWhiteSpace(primaryNameAttribute))
            {
                EntityMetadata meta = MetadataHelper.GetEntityMetadata(orgService, entityType, EntityFilters.Entity);
                primaryNameAttribute = meta.PrimaryNameAttribute;
            }
            QueryExpression query = new QueryExpression(entityType) { ColumnSet = columnSet };
            query.Criteria.AddCondition(primaryNameAttribute, ConditionOperator.Equal, primaryName);
            return RetrieveSingleEntity(orgService, query);
        }

        /// <summary>
        /// Returns a single record using a QueryExpression
        /// </summary>
        public static Entity RetrieveSingleEntity(IOrganizationService orgService, QueryExpression query)
        {
            query.TopCount = 1;
            EntityCollection results = orgService.RetrieveMultiple(query);
            return (results.Entities == null || results.Entities.Count < 1) ? null : results.Entities[0];
        }

        /// <summary>
        /// Retrieves an Entity suppressing the Exception raised if it does not exist
        /// </summary>
        public static Entity TryRetrieve(IOrganizationService orgService, string entityType, Guid entityId, ColumnSet columnSet)
        {
            try
            {
                return orgService.Retrieve(entityType, entityId, columnSet);
            }
            catch (Exception ex)
            {
                if (!ex.Message.EndsWith("Does Not Exist") && !ex.Message.StartsWith("The 'Retrieve' method does not support entities of type")) throw ex;
            }
            return null;
        }
    }
}