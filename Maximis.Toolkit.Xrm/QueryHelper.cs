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
        private static ColumnSet emptyColSet = new ColumnSet();

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

        public static Entity RetrieveEntityByName(CrmContext context, string entityType, string primaryNameValue, ColumnSet colSet = null)
        {
            EntityMetadata meta = MetadataHelper.GetEntityMetadata(context, entityType);
            QueryExpression query = new QueryExpression(entityType);
            if (colSet != null) query.ColumnSet = colSet;
            query.Criteria.AddCondition(meta.PrimaryNameAttribute, ConditionOperator.Equal, primaryNameValue);
            return RetrieveSingleEntity(context.OrganizationService, query);
        }

        public static EntityCollection RetrieveManyToManyRelatedEntities(IOrganizationService orgService, EntityReference entityRef, ManyToManyRelationshipMetadata relManyMany, ColumnSet columns = null, OrderExpression order = null)
        {
            // Create Query
            QueryExpression manyToManyQuery = new QueryExpression();
            if (columns != null) manyToManyQuery.ColumnSet = columns;
            if (order != null) manyToManyQuery.Orders.Add(order);

            // Configure Query to return the "other type" of Entity from that of "entityRef"
            if (entityRef.LogicalName == relManyMany.Entity1LogicalName)
            {
                manyToManyQuery.EntityName = relManyMany.Entity2LogicalName;
                LinkEntity le = manyToManyQuery.AddLink(relManyMany.IntersectEntityName, relManyMany.Entity2IntersectAttribute, relManyMany.Entity2IntersectAttribute);
                le.LinkCriteria.AddCondition(relManyMany.Entity1IntersectAttribute, ConditionOperator.Equal, entityRef.Id);
            }
            else
            {
                manyToManyQuery.EntityName = relManyMany.Entity1LogicalName;
                LinkEntity le = manyToManyQuery.AddLink(relManyMany.IntersectEntityName, relManyMany.Entity1IntersectAttribute, relManyMany.Entity1IntersectAttribute);
                le.LinkCriteria.AddCondition(relManyMany.Entity2IntersectAttribute, ConditionOperator.Equal, entityRef.Id);
            }

            // Return results of query
            return orgService.RetrieveMultiple(manyToManyQuery);
        }

        public static EntityCollection RetrieveOneToManyReferencedEntities(IOrganizationService orgService, EntityReference entityRef, OneToManyRelationshipMetadata relOneMany, ColumnSet columns = null, OrderExpression order = null)
        {
            if (columns == null) columns = emptyColSet;

            RetrieveRequest req = new RetrieveRequest
            {
                Target = entityRef,
                ColumnSet = emptyColSet,
                RelatedEntitiesQuery = new RelationshipQueryCollection()
            };
            QueryExpression relatedQuery = new QueryExpression(relOneMany.ReferencedEntity);
            if (columns != null) relatedQuery.ColumnSet = columns;
            if (order != null) relatedQuery.Orders.Add(order);
            req.RelatedEntitiesQuery.Add(new Relationship(relOneMany.SchemaName) { PrimaryEntityRole = EntityRole.Referencing }, relatedQuery);
            RelatedEntityCollection result = ((RetrieveResponse)orgService.Execute(req)).Entity.RelatedEntities;
            return (result == null || result.Count == 0) ? null : result.First().Value;
        }

        public static EntityCollection RetrieveOneToManyReferencingEntities(IOrganizationService orgService, EntityReference entityRef, OneToManyRelationshipMetadata relOneMany, ColumnSet columns = null, OrderExpression order = null)
        {
            if (columns == null) columns = emptyColSet;
            QueryExpression referencingQuery = new QueryExpression(relOneMany.ReferencingEntity);
            if (columns != null) referencingQuery.ColumnSet = columns;
            if (order != null) referencingQuery.Orders.Add(order);
            referencingQuery.Criteria.AddCondition(relOneMany.ReferencingAttribute, ConditionOperator.Equal, entityRef.Id);
            return orgService.RetrieveMultiple(referencingQuery);
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
        /// Returns all records related to "entityReference" by the relationship defined in
        /// "relationshipMetadata". Use MetadataHelper.GetRelationshipMetadata to get "relationshipMetadata".
        /// </summary>
        public static EntityCollection RetrieveRelatedEntities(IOrganizationService orgService, EntityReference entityRef, RelationshipMetadataBase relationshipMetadata, ColumnSet columns = null, OrderExpression orders = null)
        {
            if (columns == null) columns = emptyColSet;

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
                    if (entityRef.LogicalName == relOneMany.ReferencedEntity)
                    {
                        // If "entity" is the Referenced Entity, return all Referencing
                        // Entities which have a Lookup set to "entityReference"
                        return RetrieveOneToManyReferencingEntities(orgService, entityRef, relOneMany, columns);
                    }
                    else
                    {
                        // If "entity" is the Referencing Entity, return all Referenced entities
                        return RetrieveOneToManyReferencedEntities(orgService, entityRef, relOneMany, columns);
                    }

                // Many to Many
                case RelationshipType.ManyToManyRelationship:

                    // Many to Many relationships are really 2x One to Many with an "Intersect
                    // Entity" in between: A -< X >- B. If "entity" is an A, we want to
                    // return all B's related to all X's related to A. If "entity" is a B,
                    // we want to return all A's related to all X's related to B.

                    // Cast Metadata
                    ManyToManyRelationshipMetadata relManyMany = relationshipMetadata as ManyToManyRelationshipMetadata;

                    return RetrieveManyToManyRelatedEntities(orgService, entityRef, relManyMany, columns);
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