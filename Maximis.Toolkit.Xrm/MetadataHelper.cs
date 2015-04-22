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
        public static string ConvertOptionSetValueToString(EnumAttributeMetadata attributeMetadata, Entity entity)
        {
            return ConvertOptionSetValueToString(attributeMetadata, entity.GetAttributeValue<OptionSetValue>(attributeMetadata.LogicalName));
        }

        /// <summary>
        /// Looks up an OptionSetValue which has the supplied label
        /// </summary>
        public static OptionSetValue ConvertStringToOptionSetValue(EnumAttributeMetadata attributeMetadata,
            string labelText)
        {
            OptionMetadata optionMetadata =
                attributeMetadata.OptionSet.Options.SingleOrDefault(q => q.Label.UserLocalizedLabel.Label == labelText);
            return optionMetadata == null ? null : new OptionSetValue((int)optionMetadata.Value);
        }

        /// <summary>
        /// Returns the Metadata for all Entities
        /// </summary>
        public static EntityMetadata[] GetAllEntityMetadata(IOrganizationService orgService,
              EntityFilters filters = (EntityFilters.Entity | EntityFilters.Attributes))
        {
            return
                 ((RetrieveAllEntitiesResponse)
                     orgService.Execute(new RetrieveAllEntitiesRequest
                     {
                         EntityFilters = filters
                     })).EntityMetadata;
        }

        /// <summary>
        /// Returns the Metadata for an Attribute
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

        /// <summary>
        /// Converts attribute values of different types to a string
        /// </summary>
        public static string GetAttributeValueAsDisplayString(IOrganizationService orgService, MetadataCache metaCache, Entity entity, string attributeName, DisplayStringOptions options = null)
        {
            // Return null if value not present
            if (entity == null || !entity.Contains(attributeName) || entity[attributeName] == null)
            {
                return null;
            };

            // Default Display String Options if required
            if (options == null) options = new DisplayStringOptions();

            // Get value as object
            object attributeValue = entity[attributeName];

            // Get Attribute Metadata (handling if value is an AliasedValue, from a LinkedEntity
            AliasedValue aliasedValue = attributeValue as AliasedValue;
            AttributeMetadata attributeMetadata = null;
            if (aliasedValue != null)
            {
                attributeMetadata = GetAttributeMetadata(metaCache.GetEntityMetadata(orgService, aliasedValue.EntityLogicalName), aliasedValue.AttributeLogicalName);
                attributeValue = aliasedValue.Value;
            }
            else
            {
                attributeMetadata = GetAttributeMetadata(metaCache.GetEntityMetadata(orgService, entity.LogicalName), attributeName);
            }

            // Output depending on Attribute Type
            switch (attributeMetadata.AttributeType)
            {
                case AttributeTypeCode.Picklist:
                case AttributeTypeCode.State:
                case AttributeTypeCode.Status:
                    OptionSetValue optionSetVal = attributeValue as OptionSetValue;
                    string optionSetStringVal = ConvertOptionSetValueToString((EnumAttributeMetadata)attributeMetadata, optionSetVal);
                    return string.IsNullOrEmpty(options.OptionSetFormat) ? optionSetStringVal : string.Format(options.OptionSetFormat, optionSetVal.Value, optionSetStringVal);

                case AttributeTypeCode.Lookup:
                case AttributeTypeCode.Owner:
                case AttributeTypeCode.Customer:
                    EntityReference eRef = attributeValue as EntityReference;
                    if (eRef == null) return null;
                    if (string.IsNullOrEmpty(eRef.Name)) return GetEntityReferenceName(orgService, metaCache, eRef.LogicalName, eRef.Id);
                    return string.IsNullOrEmpty(options.LookupFormat) ? eRef.Name : string.Format(options.LookupFormat, eRef.Id, eRef.LogicalName, eRef.Name);

                case AttributeTypeCode.DateTime:
                    DateTime dateTime = (DateTime)attributeValue;
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
                    return (bool)attributeValue ? options.BoolTrue : options.BoolFalse;

                case AttributeTypeCode.Money:
                    return ((Money)attributeValue).Value.ToString("#.00");

                default:
                    string result = attributeValue.ToString();
                    if (options.CleanWhiteSpace && !string.IsNullOrEmpty(result))
                    {
                        result = REGEX_WHITESPACE.Replace(result, " ").Trim();
                    }
                    return result;
            }
        }

        /// <summary>
        /// Returns the Metadata for an Entity
        /// </summary>
        public static EntityMetadata GetEntityMetadata(IOrganizationService orgService, string entityType,
            EntityFilters filters = (EntityFilters.Entity | EntityFilters.Attributes))
        {
            return ((RetrieveEntityResponse)
            orgService.Execute(new RetrieveEntityRequest
            {
                EntityFilters = filters,
                LogicalName = entityType
            })).EntityMetadata;
        }

        /// <summary>
        /// Returns a Dictionary of Option Set values and text
        /// </summary>
        public static Dictionary<int, string> GetOptionSetLookup(EnumAttributeMetadata attributeMetadata)
        {
            return attributeMetadata.OptionSet.Options.OrderBy(o => o.Value.Value).ToDictionary(k => k.Value.Value,
                v => v.Label.UserLocalizedLabel.Label);
        }

        public static List<Entity> GetRelationshipAttributeMappings(IOrganizationService orgService, string primaryEntityType, string relatedEntityType)
        {
            QueryExpression query = new QueryExpression("attributemap") { ColumnSet = new ColumnSet("sourceattributename", "targetattributename", "entitymapid") };
            LinkEntity le = query.AddLink("entitymap", "entitymapid", "entitymapid");
            le.LinkCriteria.AddCondition("sourceentityname", ConditionOperator.Equal, primaryEntityType);
            le.LinkCriteria.AddCondition("targetentityname", ConditionOperator.Equal, relatedEntityType);

            return QueryHelper.RetrieveAllEntities(orgService, query);
        }

        /// <summary>
        /// Returns the Metadata for a Relationship
        /// </summary>
        public static RelationshipMetadataBase GetRelationshipMetadata(IOrganizationService orgService, string relationshipName)
        {
            return
                ((RetrieveRelationshipResponse)orgService.Execute(new RetrieveRelationshipRequest { Name = relationshipName }))
                    .RelationshipMetadata;
        }

        private static string GetEntityReferenceName(IOrganizationService orgService, MetadataCache metaCache, string entityType, Guid id)
        {
            // Special case if Id is Guid.Empty
            if (id == Guid.Empty) return "Guid.Empty";

            // See if Entity Reference has already been retrieved
            string cacheKey = string.Format("{0}{1:N}", entityType, id);
            string cachedName = metaCache.RetrieveValue(cacheKey) as string;
            if (!string.IsNullOrEmpty(cachedName)) return cachedName;

            // If not, retrieve record
            EntityMetadata entityMeta = metaCache.GetEntityMetadata(orgService, entityType, EntityFilters.Entity);
            Entity e = orgService.Retrieve(entityType, id, new ColumnSet(entityMeta.PrimaryNameAttribute));

            // Populate result
            string result = "Null Entity";
            if (e != null)
            {
                result = e.GetAttributeValue<string>(entityMeta.PrimaryNameAttribute);
                if (string.IsNullOrEmpty(result)) result = (entityMeta.PrimaryNameAttribute + " is empty");
            }

            // Store in case it is encountered again
            metaCache.StoreValue(cacheKey, result);
            return result;
        }
    }
}