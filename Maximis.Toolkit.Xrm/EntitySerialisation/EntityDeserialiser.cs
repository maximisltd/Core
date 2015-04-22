using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace Maximis.Toolkit.Xrm.EntitySerialisation
{
    public class EntityDeserialiser
    {
        private List<EntityReference> lookupCache = new List<EntityReference>();
        private MetadataCache metaCache;

        public EntityDeserialiser()
        {
            this.metaCache = new MetadataCache();
        }

        public EntityDeserialiser(MetadataCache metaCache)
        {
            this.metaCache = metaCache;
        }

        public Entity DeserialiseEntity(IOrganizationService orgService, string entityXml, DeserialisationOptions options = null)
        {
            // Default options
            if (options == null) options = new DeserialisationOptions();

            // Load XML
            XmlDocument xd = new XmlDocument();
            xd.LoadXml(entityXml);
            if (xd.DocumentElement.Name != "ent") return null;

            // Get Metadata for Entity Type
            string logicalName = xd.DocumentElement.GetAttribute("ln");
            EntityMetadata entityMeta = metaCache.GetEntityMetadata(orgService, logicalName);
            if (entityMeta == null) return null;

            // Create Entity
            Entity result = new Entity(logicalName);

            // Set ID if we have it
            string idAttr = xd.DocumentElement.GetAttribute("id");
            if (!string.IsNullOrEmpty(idAttr)) result.Id = idAttr.ToGuid();

            // Loop through each Attribute
            foreach (XmlElement attrNode in xd.DocumentElement.SelectNodes("attr"))
            {
                string attrName = attrNode.GetAttribute("ln");
                if (string.IsNullOrEmpty(attrName)) continue;

                AttributeMetadata attrMeta = MetadataHelper.GetAttributeMetadata(entityMeta, attrName);

                if (attrMeta == null)
                {
                    if (options.IgnoreUnknownAttributes) continue;
                    else throw new Exception(string.Format("Unknown Attribute: '{0}'", attrName));
                }

                if (string.IsNullOrEmpty(attrNode.InnerText))
                {
                    if (options.SetToNullIfEmpty != null && options.SetToNullIfEmpty.Contains(attrMeta.LogicalName))
                    {
                        result[attrMeta.LogicalName] = null;
                    }
                    continue;
                }

                switch (attrMeta.AttributeType)
                {
                    case AttributeTypeCode.Uniqueidentifier:
                        result[attrMeta.LogicalName] = new Guid(attrNode.InnerText);
                        break;

                    case AttributeTypeCode.Picklist:
                    case AttributeTypeCode.State:
                    case AttributeTypeCode.Status:
                        string valAttr = attrNode.GetAttribute("val");
                        if (string.IsNullOrEmpty(valAttr))
                        {
                            result[attrMeta.LogicalName] = MetadataHelper.ConvertStringToOptionSetValue((EnumAttributeMetadata)attrMeta, attrNode.InnerText);
                        }
                        else
                        {
                            result[attrMeta.LogicalName] = new OptionSetValue(valAttr.ToInt());
                        }
                        break;

                    case AttributeTypeCode.Lookup:
                    case AttributeTypeCode.Owner:
                    case AttributeTypeCode.Customer:
                        EntityReference entityRef = GetEntityReference(orgService, attrNode, (LookupAttributeMetadata)attrMeta);
                        if (entityRef != null) result[attrMeta.LogicalName] = entityRef;
                        break;

                    case AttributeTypeCode.DateTime:
                        result[attrMeta.LogicalName] = attrNode.InnerText.ToDateTime();
                        break;

                    case AttributeTypeCode.Boolean:
                        result[attrMeta.LogicalName] = attrNode.InnerText.ToBoolean();
                        break;

                    case AttributeTypeCode.Decimal:
                        result[attrMeta.LogicalName] = attrNode.InnerText.ToDecimal();
                        break;

                    case AttributeTypeCode.Money:
                        result[attrMeta.LogicalName] = new Money(attrNode.InnerText.ToDecimal());
                        break;

                    case AttributeTypeCode.Integer:
                        result[attrMeta.LogicalName] = attrNode.InnerText.ToInt();
                        break;

                    case AttributeTypeCode.Memo:
                        int maxLengthMemo = ((MemoAttributeMetadata)attrMeta).MaxLength.Value;
                        if (attrNode.InnerText.Length > maxLengthMemo)
                            result[attrMeta.LogicalName] = attrNode.InnerText.Substring(0, maxLengthMemo);
                        else
                            result[attrMeta.LogicalName] = attrNode.InnerText;
                        break;

                    case AttributeTypeCode.String:
                        int maxLengthString = ((StringAttributeMetadata)attrMeta).MaxLength.Value;
                        if (attrNode.InnerText.Length > maxLengthString)
                            result[attrMeta.LogicalName] = attrNode.InnerText.Substring(0, maxLengthString);
                        else
                            result[attrMeta.LogicalName] = attrNode.InnerText;
                        break;

                    case AttributeTypeCode.PartyList:
                        EntityCollection partyList = new EntityCollection();
                        foreach (XmlNode partyNode in attrNode.SelectNodes("ent"))
                        {
                            partyList.Entities.Add(DeserialiseEntity(orgService, partyNode.OuterXml, options));
                        }
                        result[attrMeta.LogicalName] = partyList;
                        break;

                    case AttributeTypeCode.EntityName:
                        break;

                    default:
                        throw new NotSupportedException(string.Format("Entity deserialiser: unsupported attribute type '{0}' (entity '{1}', attribute '{2}')", attrMeta.AttributeType, entityMeta.LogicalName, attrMeta.LogicalName));
                }
            }

            return result;
        }

        private EntityReference GetEntityReference(IOrganizationService orgService, XmlElement attrNode, LookupAttributeMetadata attrMeta)
        {
            List<EntityReference> possibles = new List<EntityReference>();

            // First, see if there is an <ent> element within this <attr> element
            XmlElement entityNode = attrNode.SelectSingleNode("ent") as XmlElement;

            // If not, assume Inner Text is Primary Name Value, and reference could be any of the types supported by this attribute
            if (entityNode == null)
            {
                string name = attrNode.InnerText;
                if (string.IsNullOrWhiteSpace(name)) return null;
                foreach (string lookupEntityType in attrMeta.Targets)
                {
                    possibles.Add(new EntityReference(lookupEntityType, Guid.Empty) { Name = name });
                }
            }
            else
            {
                string lookupEntityType = entityNode.GetAttribute("ln");
                if (lookupEntityType == "adx_webfile") return null; // TEMP
                EntityMetadata lookupMeta = metaCache.GetEntityMetadata(orgService, lookupEntityType);
                XmlElement primaryNode = entityNode.SelectSingleNode("attr[@ln='" + lookupMeta.PrimaryNameAttribute + "']") as XmlElement;
                possibles.Add(new EntityReference(lookupEntityType, entityNode.GetAttribute("id").ToGuid()) { Name = (primaryNode == null) ? null : primaryNode.InnerText });
            }

            foreach (EntityReference possible in possibles)
            {
                if (possible.Id == Guid.Empty && string.IsNullOrWhiteSpace(possible.Name)) continue;

                // See if "possible" already exists in Lookup Cache - it will be in the cache with an empty GUID if it's already been checked and doesn't exist in CRM.
                EntityReference result = null;
                if (possible.Id != Guid.Empty) result = lookupCache.SingleOrDefault(q => q.LogicalName == possible.LogicalName && q.Id == possible.Id);
                if (result == null && !string.IsNullOrEmpty(possible.Name)) result = lookupCache.SingleOrDefault(q => q.LogicalName == possible.LogicalName && q.Name == possible.Name);

                // If not, attempt to retrieve it
                if (result == null)
                {
                    EntityMetadata lookupMeta = metaCache.GetEntityMetadata(orgService, possible.LogicalName);

                    QueryExpression query = new QueryExpression(lookupMeta.LogicalName);
                    query.Criteria.FilterOperator = LogicalOperator.Or;
                    if (possible.Id != Guid.Empty) query.Criteria.AddCondition(lookupMeta.PrimaryIdAttribute, ConditionOperator.Equal, possible.Id);
                    if (!string.IsNullOrWhiteSpace(possible.Name)) query.Criteria.AddCondition(lookupMeta.PrimaryNameAttribute, ConditionOperator.Equal, possible.Name);

                    if (lookupMeta.LogicalName == "systemuser")
                    {
                        FilterExpression fe = new FilterExpression();
                        fe.AddCondition("isdisabled", ConditionOperator.Equal, false);
                        fe.AddCondition("accessmode", ConditionOperator.Equal, 0);
                        query.Criteria.AddFilter(fe);
                    }

                    Entity lookupEntity = QueryHelper.RetrieveSingleEntity(orgService, query);

                    if (lookupEntity == null) result = new EntityReference(lookupMeta.LogicalName, Guid.Empty);
                    else result = lookupEntity.ToEntityReference();
                    result.Name = possible.Name;

                    lookupCache.Add(result);
                }

                if (result.Id != Guid.Empty) return result;
            }
            return null;
        }
    }
}