using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Maximis.Toolkit.Xrm
{
    public static class FetchXmlHelper
    {
        public static string GetFetchXmlFromQueryExpression(IOrganizationService orgService, QueryExpression query)
        {
            return ((QueryExpressionToFetchXmlResponse)orgService.Execute(new QueryExpressionToFetchXmlRequest { Query = query })).FetchXml;
        }

        public static QueryExpression GetQueryExpressionFromFetchXml(IOrganizationService orgService, string fetchXml)
        {
            return ((FetchXmlToQueryExpressionResponse)orgService.Execute(new FetchXmlToQueryExpressionRequest { FetchXml = fetchXml })).Query;
        }

        public static List<Entity> PerformFetchXmlQuery(IOrganizationService orgService, string fetchXml, bool mergeLinkedEntities = false)
        {
            // Perform FetchXML query
            EntityCollection entities = orgService.RetrieveMultiple(new FetchExpression(fetchXml));
            if (entities.Entities.Count == 0) return null;

            List<Entity> result = entities.Entities.ToList();

            if (mergeLinkedEntities)
            {
                List<Entity> relatedEntities = new List<Entity>();
                foreach (Entity mainEntity in result)
                {
                    relatedEntities.AddRange(ExtractRelatedEntities(orgService, mainEntity));
                }
                result.AddRange(relatedEntities);
            }

            return result;
        }

        private static List<Entity> ExtractRelatedEntities(IOrganizationService orgService, Entity mainEntity)
        {
            Dictionary<string, Entity> relatedEntities = new Dictionary<string, Entity>();

            foreach (string fullAttributeName in mainEntity.Attributes.Keys)
            {
                if (fullAttributeName.Contains('.') && mainEntity.HasAttributeWithValue(fullAttributeName))
                {
                    AliasedValue aliasedValue = mainEntity.Attributes[fullAttributeName] as AliasedValue;
                    if (aliasedValue == null) continue;

                    string entityKey = fullAttributeName.LeftOfFirst('.');

                    if (!relatedEntities.ContainsKey(entityKey))
                    {
                        relatedEntities[entityKey] = new Entity(aliasedValue.EntityLogicalName);
                    }
                    if (aliasedValue.AttributeLogicalName == aliasedValue.EntityLogicalName + "id")
                    {
                        relatedEntities[entityKey].Id = (Guid)aliasedValue.Value;
                    }
                    relatedEntities[entityKey][aliasedValue.AttributeLogicalName] = aliasedValue.Value;
                }
            }

            return relatedEntities.Values.ToList();
        }
    }
}