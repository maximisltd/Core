using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm
{
    public static class ExtensionMethods
    {
        /// <summary>
        /// Returns a DateTime attribute value as the local time
        /// </summary>
        public static DateTime GetLocalDateTimeValue(this Entity entity, string attributeName)
        {
            DateTime result = entity.GetAttributeValue<DateTime>(attributeName);
            if (result == DateTime.MinValue) return DateTime.MinValue;
            return result.ToLocalTime();
        }

        /// <summary>
        /// Returns true if an attribute is present on an entity and it contains a value
        /// </summary>
        public static bool HasAttributeWithValue(this Entity entity, string attributeName)
        {
            return entity != null && entity.Contains(attributeName) && entity[attributeName] != null;
        }

        /// <summary>
        /// Returns true if an attribute is present on an entity and is null
        /// </summary>
        public static bool HasNullAttribute(this Entity entity, string attributeName)
        {
            return entity.Contains(attributeName) && entity[attributeName] == null;
        }

        /// <summary>
        /// Copies attributes present in the secondary, but not present in the primary, into the primary
        /// </summary>
        public static Entity Merge(this Entity primary, Entity secondary)
        {
            if (primary == null && secondary == null) return null;
            if (primary == null) return secondary;
            if (secondary == null) return primary;
            foreach (KeyValuePair<string, object> attribute in secondary.Attributes)
            {
                if (!primary.Contains(attribute.Key)) primary[attribute.Key] = attribute.Value;
            }
            return primary;
        }

        /// <summary>
        /// Retrieves an Entity from an EntityReference
        /// </summary>
        public static Entity RetrieveEntity(this EntityReference entityRef, IOrganizationService orgService, ColumnSet columns)
        {
            if (entityRef == null) return null;
            if (orgService == null) return null;
            return orgService.Retrieve(entityRef.LogicalName, entityRef.Id, columns == null ? new ColumnSet() : columns);
        }

        /// <summary>
        /// Converts an EntityReference into an Entity
        /// </summary>
        public static Entity ToEntity(this EntityReference entityRef)
        {
            return new Entity(entityRef.LogicalName) { Id = entityRef.Id };
        }

        /// <summary>
        /// Converts an Entity into an EntityReference
        /// </summary>
        public static EntityReference ToEntityReference(this Entity entity)
        {
            return new EntityReference(entity.LogicalName, entity.Id);
        }
    }
}