using Maximis.Toolkit.Caching;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm
{
    public class MetadataCache
    {
        private Dictionary<string, EntityMetadata> metadataCache;
        private CacheHelper<object> valueStore;

        public EntityMetadata GetEntityMetadata(IOrganizationService orgService, string entityType, EntityFilters filters = (EntityFilters.Entity | EntityFilters.Attributes))
        {
            if (metadataCache == null) metadataCache = new Dictionary<string, EntityMetadata>();
            if (!metadataCache.ContainsKey(entityType))
            {
                metadataCache[entityType] = ((RetrieveEntityResponse)
                    orgService.Execute(new RetrieveEntityRequest
                    {
                        EntityFilters = filters,
                        LogicalName = entityType
                    })).EntityMetadata;
            }
            return metadataCache[entityType];
        }

        public object RetrieveValue(string key)
        {
            if (valueStore == null) return null;
            return valueStore.Dictionary.ContainsKey(key) ? valueStore.Dictionary[key] : null;
        }

        public void StoreValue(string key, object val)
        {
            if (valueStore == null) valueStore = new CacheHelper<object>(3);
            valueStore.Dictionary[key] = val;
        }
    }
}