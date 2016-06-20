using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Maximis.Toolkit.Xrm
{
    public static class MetadataHelper
    {
        private static readonly Regex REGEX_WHITESPACE = new Regex("\\s+", RegexOptions.Compiled);

        #region Option Set

        /// <summary>
        /// Returns the Label of the selected option in an OptionSet or similar field
        /// </summary>
        public static string ConvertOptionSetValueToString(CrmContext context, string entityName, string attributeName, OptionSetValue optionSetValue)
        {
            EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, entityName);
            EnumAttributeMetadata attrMeta = (EnumAttributeMetadata)entityMeta.Attributes.Single(q => q.LogicalName == attributeName);
            return ConvertOptionSetValueToString(attrMeta, optionSetValue);
        }

        /// <summary>
        /// Returns the Label of the selected option in an OptionSet or similar field
        /// </summary>
        public static string ConvertOptionSetValueToString(EnumAttributeMetadata attributeMetadata, OptionSetValue optionSetValue)
        {
            OptionMetadata optionMetadata =
                attributeMetadata.OptionSet.Options.SingleOrDefault(q => q.Value == optionSetValue.Value);
            return optionMetadata == null ? null : optionMetadata.Label.UserLocalizedLabel.Label;
        }

        /// <summary>
        /// Returns the Label of the selected option in an OptionSet or similar field
        /// </summary>
        public static string ConvertOptionSetValueToString(CrmContext context, Entity entity, string attributeName)
        {
            return ConvertOptionSetValueToString(context, entity.LogicalName, attributeName, entity.GetAttributeValue<OptionSetValue>(attributeName));
        }

        /// <summary>
        /// Looks up an OptionSetValue which has the supplied label
        /// </summary>
        public static OptionSetValue ConvertStringToOptionSetValue(CrmContext context, string entityName, string attributeName, string labelText)
        {
            EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, entityName);
            EnumAttributeMetadata attrMeta = (EnumAttributeMetadata)entityMeta.Attributes.Single(q => q.LogicalName == attributeName);
            return ConvertStringToOptionSetValue(attrMeta, labelText);
        }

        /// <summary>
        /// Looks up an OptionSetValue which has the supplied label
        /// </summary>
        public static OptionSetValue ConvertStringToOptionSetValue(EnumAttributeMetadata attributeMetadata,
            string labelText)
        {
            OptionMetadata optionMetadata =
                attributeMetadata.OptionSet.Options.SingleOrDefault(q => q.Label.UserLocalizedLabel.Label.ToUpper() == labelText.ToUpper());
            return optionMetadata == null ? null : new OptionSetValue((int)optionMetadata.Value);
        }

        /// <summary>
        /// Returns a Dictionary of Option Set values and text
        /// </summary>
        public static Dictionary<int, string> GetOptionSetLookup(EnumAttributeMetadata attributeMetadata)
        {
            return attributeMetadata.OptionSet.Options.OrderBy(o => o.Value.Value).ToDictionary(k => k.Value.Value,
                v => v.Label.UserLocalizedLabel.Label);
        }

        #endregion Option Set

        #region Attribute Values

        /// <summary>
        /// Converts attribute values of different types to a string
        /// </summary>
        public static string GetAttributeValueAsDisplayString(CrmContext context, Entity entity, EntityMetadata entityMetadata, string attributeName, DisplayStringOptions options = null)
        {
            AttributeMetadata attrMeta = entityMetadata.Attributes.Single(q => q.LogicalName == attributeName);
            return GetAttributeValueAsDisplayString(context, entity, attrMeta, options);
        }

        /// <summary>
        /// Converts attribute values of different types to a string
        /// </summary>
        public static string GetAttributeValueAsDisplayString(CrmContext context, Entity entity, AttributeMetadata attributeMetadata, DisplayStringOptions options = null)
        {
            // Return null if value not present
            string attributeName = attributeMetadata.LogicalName;
            if (entity == null || !entity.Contains(attributeName) || entity[attributeName] == null)
            {
                return null;
            };

            return GetAttributeValueAsDisplayString(context, entity[attributeName], attributeMetadata, options);
        }

        /// <summary>
        /// Converts attribute values of different types to a string
        /// </summary>
        public static string GetAttributeValueAsDisplayString(CrmContext context, Entity entity, string attributeName, DisplayStringOptions options = null)
        {
            EntityMetadata entityMeta = GetEntityMetadata(context, entity.LogicalName);
            AttributeMetadata attrMeta = GetAttributeMetadata(entityMeta, attributeName);
            return GetAttributeValueAsDisplayString(context, entity, attrMeta, options);
        }

        public static string GetAttributeValueAsDisplayString(CrmContext context, object attributeValue, AttributeMetadata attributeMetadata, DisplayStringOptions options = null)
        {
            // Create Default DisplayStringOptions if not supplied
            if (options == null) options = new DisplayStringOptions();

            // Output depending on Attribute Type
            switch (attributeMetadata.AttributeType)
            {
                case AttributeTypeCode.Picklist:
                case AttributeTypeCode.State:
                case AttributeTypeCode.Status:
                    if (attributeValue == null) return null;
                    OptionSetValue optionSetVal = attributeValue as OptionSetValue;
                    string optionSetStringVal = ConvertOptionSetValueToString((EnumAttributeMetadata)attributeMetadata, optionSetVal);
                    return string.IsNullOrEmpty(options.OptionSetFormat) ? optionSetStringVal : string.Format(options.OptionSetFormat, optionSetVal.Value, optionSetStringVal);

                case AttributeTypeCode.Lookup:
                case AttributeTypeCode.Owner:
                case AttributeTypeCode.Customer:
                    if (attributeValue == null) return null;
                    EntityReference eRef = attributeValue as EntityReference;
                    if (eRef == null) return null;
                    if (string.IsNullOrEmpty(eRef.Name)) return GetPrimaryNameValue(context, eRef.LogicalName, eRef.Id);
                    return string.IsNullOrEmpty(options.LookupFormat) ? eRef.Name : string.Format(options.LookupFormat, eRef.Id, eRef.LogicalName, eRef.Name);

                case AttributeTypeCode.DateTime:
                    if (attributeValue == null) return null;
                    DateTime dateTime = ((DateTime)attributeValue).ToLocalTime();
                    if (string.IsNullOrEmpty(options.DateFormat))
                    {
                        if (dateTime.Hour == 0 && dateTime.Minute == 0 && dateTime.Second == 0)
                        {
                            return dateTime.ToShortDateString();
                        }
                        else
                        {
                            return dateTime.ToString();
                        }
                    }
                    return dateTime.ToString(options.DateFormat);

                case AttributeTypeCode.Boolean:
                    if (attributeValue == null) return options.BoolFalse;
                    else return (bool)attributeValue ? options.BoolTrue : options.BoolFalse;

                case AttributeTypeCode.Money:
                    if (attributeValue == null) return 0.ToString("C");
                    else return ((Money)attributeValue).Value.ToString("C");

                default:
                    if (attributeValue == null) return null;
                    string result = attributeValue.ToString();
                    if (options.CleanWhiteSpace && !string.IsNullOrEmpty(result))
                    {
                        result = REGEX_WHITESPACE.Replace(result, " ").Trim();
                    }
                    return result;
            }
        }

        #endregion Attribute Values

        #region Entity Metadata

        /// <summary>
        /// Returns the Metadata for all Entities
        /// </summary>
        public static EntityMetadata[] GetAllEntityMetadata(IOrganizationService orgService,
              EntityFilters filters = (EntityFilters.Entity | EntityFilters.Attributes))
        {
            return ((RetrieveAllEntitiesResponse)orgService.Execute(new RetrieveAllEntitiesRequest { EntityFilters = filters })).EntityMetadata;
        }

        /// <summary>
        /// Returns the Metadata for an Entity
        /// </summary>
        public static EntityMetadata GetEntityMetadata(IOrganizationService orgService, string entityType, EntityFilters filters = (EntityFilters.Entity | EntityFilters.Attributes))
        {
            return ((RetrieveEntityResponse)orgService.Execute(new RetrieveEntityRequest
            {
                EntityFilters = filters,
                LogicalName = entityType
            })).EntityMetadata;
        }

        /// <summary>
        /// Returns the Metadata for an Entity
        /// </summary>
        public static EntityMetadata GetEntityMetadata(CrmContext context, string entityType, EntityFilters filters = (EntityFilters.Entity | EntityFilters.Attributes))
        {
            string key = string.Format("{0}_{1}", entityType, filters);
            EntityMetadata result = context.CacheManager.Get<EntityMetadata>(key);
            if (result == null)
            {
                result = GetEntityMetadata(context.OrganizationService, entityType, filters);
                context.CacheManager.Set<EntityMetadata>(key, result);
            }
            return result;
        }

        #endregion Entity Metadata

        #region Attribute Metadata

        /// <summary>
        /// Returns the Metadata for an Attribute (with caching)
        /// </summary>
        public static AttributeMetadata GetAttributeMetadata(CrmContext context, string entityType,
            string attributeName)
        {
            EntityMetadata entityMeta = GetEntityMetadata(context, entityType);
            return GetAttributeMetadata(entityMeta, attributeName);
        }

        /// <summary>
        /// Returns the Metadata for an Attribute (without caching)
        /// </summary>
        public static AttributeMetadata GetAttributeMetadata(IOrganizationService orgService, string entityType,
            string attributeName)
        {
            return
            ((RetrieveAttributeResponse)
                orgService.Execute(new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityType,
                    LogicalName = attributeName
                })).AttributeMetadata;
        }

        /// <summary>
        /// Returns the Metadata for an Attribute from already-retrieved Entity Metadata
        /// </summary>
        public static AttributeMetadata GetAttributeMetadata(EntityMetadata entityMetadata, string attributeName)
        {
            return entityMetadata.Attributes.SingleOrDefault(q => q.LogicalName == attributeName);
        }

        #endregion Attribute Metadata

        #region Relationships

        public static List<RelationshipMetadataBase> GetAllRelationships(IOrganizationService orgService, string entityType, params string[] relatedTypes)
        {
            return GetAllRelationships(GetEntityMetadata(orgService, entityType, EntityFilters.Relationships), relatedTypes);
        }

        public static List<RelationshipMetadataBase> GetAllRelationships(EntityMetadata entityMetadata, params string[] relatedTypes)
        {
            List<RelationshipMetadataBase> relationships = new List<RelationshipMetadataBase>();
            foreach (string relatedType in relatedTypes)
            {
                relationships.AddRange(entityMetadata.ManyToManyRelationships.Where(q => q.Entity1LogicalName == entityMetadata.LogicalName && q.Entity2LogicalName == relatedType));
                relationships.AddRange(entityMetadata.ManyToManyRelationships.Where(q => q.Entity2LogicalName == entityMetadata.LogicalName && q.Entity1LogicalName == relatedType));
                relationships.AddRange(entityMetadata.ManyToOneRelationships.Where(q => q.ReferencedEntity == entityMetadata.LogicalName && q.ReferencingEntity == relatedType));
                relationships.AddRange(entityMetadata.ManyToOneRelationships.Where(q => q.ReferencingEntity == entityMetadata.LogicalName && q.ReferencedEntity == relatedType));
                relationships.AddRange(entityMetadata.OneToManyRelationships.Where(q => q.ReferencedEntity == entityMetadata.LogicalName && q.ReferencingEntity == relatedType));
                relationships.AddRange(entityMetadata.OneToManyRelationships.Where(q => q.ReferencingEntity == entityMetadata.LogicalName && q.ReferencedEntity == relatedType));
            }
            return relationships;
        }

        public static List<Entity> GetRelationshipAttributeMappings(IOrganizationService orgService, string primaryEntityType, string relatedEntityType)
        {
            QueryExpression query = new QueryExpression("attributemap") { ColumnSet = new ColumnSet("sourceattributename", "targetattributename", "entitymapid") };
            LinkEntity le = query.AddLink("entitymap", "entitymapid", "entitymapid");
            le.LinkCriteria.AddCondition("sourceentityname", ConditionOperator.Equal, primaryEntityType);
            le.LinkCriteria.AddCondition("targetentityname", ConditionOperator.Equal, relatedEntityType);

            return QueryHelper.RetrieveAllEntities(orgService, query);
        }

        public static RelationshipMetadataBase GetRelationshipMetadata(CrmContext context, string entityType, string relName)
        {
            EntityMetadata entityMeta = GetEntityMetadata(context, entityType, EntityFilters.Relationships);
            RelationshipMetadataBase rel = entityMeta.OneToManyRelationships.SingleOrDefault(q => q.SchemaName == relName);
            if (rel != null) return rel;
            rel = entityMeta.ManyToOneRelationships.SingleOrDefault(q => q.SchemaName == relName);
            if (rel != null) return rel;
            rel = entityMeta.ManyToManyRelationships.SingleOrDefault(q => q.SchemaName == relName);
            return rel;
        }

        #endregion Relationships

        #region Utilities

        private static string GetPrimaryNameValue(CrmContext context, string entityType, Guid id)
        {
            // Special case if Id is Guid.Empty
            if (id == Guid.Empty) return "Guid.Empty";

            // See if Entity Reference has already been retrieved
            string cacheKey = string.Format("PrimaryName_{0}_{1:N}", entityType, id);
            string result = context.CacheManager.Get<string>(cacheKey) as string;
            if (string.IsNullOrEmpty(result))
            {
                // If not, retrieve record
                EntityMetadata entityMeta = GetEntityMetadata(context, entityType, EntityFilters.Entity);
                Entity e = context.OrganizationService.Retrieve(entityType, id, new ColumnSet(entityMeta.PrimaryNameAttribute));

                // Populate result
                result = "Null Entity";
                if (e != null)
                {
                    result = e.GetAttributeValue<string>(entityMeta.PrimaryNameAttribute);
                    if (string.IsNullOrEmpty(result)) result = (entityMeta.PrimaryNameAttribute + " is empty");
                }

                // Store in case it is encountered again
                context.CacheManager.Set<string>(cacheKey, result);
            }

            return result;
        }

        #endregion Utilities
    }
}