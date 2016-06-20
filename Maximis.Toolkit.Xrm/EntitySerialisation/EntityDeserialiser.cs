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
        private EntityReference defaultBURef;

        private Dictionary<Guid, Guid> idMappings = new Dictionary<Guid, Guid>();

        private List<EntityReference> knownExist = new List<EntityReference>();

        private List<EntityReference> knownNonExist = new List<EntityReference>();

        public EntityDeserialiser(CrmContext context)
        {
            this.context = context;
        }

        public CrmContext context { get; set; }

        /// <summary>
        /// Used to specify substitute Record IDs for lookups.
        /// If a referenced Entity cannot be found by Id, but a mapping exists between that Id and another, that will also be attempted.
        /// </summary>
        public Dictionary<Guid, Guid> IdMappings { get { return idMappings; } }

        public Entity DeserialiseEntity(string entityXml, DeserialisationOptions options = null)
        {
            // Default options
            if (options == null) options = new DeserialisationOptions();

            // Load XML
            XmlDocument xd = new XmlDocument();
            xd.LoadXml(entityXml);
            if (xd.DocumentElement.Name != "ent") return null;

            // Get Metadata for Entity Type
            string logicalName = xd.DocumentElement.GetAttribute("ln");
            EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, logicalName);
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
                    // Special Case code for Calendar Rules (and similar??)
                    if (attrNode.GetAttribute("type") == "EntityCollection")
                    {
                        List<Entity> embeddedEntities = new List<Entity>();
                        foreach (XmlElement embeddedEntityNode in attrNode.SelectNodes("ent"))
                        {
                            embeddedEntities.Add(DeserialiseEntity( embeddedEntityNode.OuterXml, options));
                        }
                        result[attrName] = new EntityCollection(embeddedEntities);
                        continue;
                    }
                    else if (options.IgnoreUnknownAttributes) continue;
                    else throw new Exception(string.Format("Unknown Attribute: '{0}'", attrName));
                }

                if (string.IsNullOrEmpty(attrNode.InnerXml))
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
                        EntityReference entityRef = GetEntityReference(context, attrNode, (LookupAttributeMetadata)attrMeta);
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

                    case AttributeTypeCode.Double:
                        result[attrMeta.LogicalName] = attrNode.InnerText.ToDouble();
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
                            partyList.Entities.Add(DeserialiseEntity( partyNode.OuterXml, options));
                        }
                        result[attrMeta.LogicalName] = partyList;
                        break;

                    case AttributeTypeCode.EntityName:
                        break;

                    default:
                        throw new NotSupportedException(string.Format("Entity deserialiser: unsupported attribute type '{0}' (entity '{1}', attribute '{2}')", attrMeta.AttributeType, entityMeta.LogicalName, attrMeta.LogicalName));
                }
            }

            // Ensure Business Owned Entities have a Business Unit
            if (entityMeta.OwnershipType == OwnershipTypes.BusinessOwned)
            {
                string buAttrName = entityMeta.Attributes.First(q => q.AttributeType == AttributeTypeCode.Lookup && ((LookupAttributeMetadata)q).Targets.Contains("businessunit")).LogicalName;
                if (!result.Contains(buAttrName))
                {
                    if (defaultBURef == null)
                    {
                        QueryExpression buQuery = new QueryExpression("businessunit");
                        buQuery.Criteria.AddCondition("parentbusinessunitid", ConditionOperator.Null);
                        defaultBURef = QueryHelper.RetrieveSingleEntity(context.OrganizationService, buQuery).ToEntityReference();
                    }

                    result[buAttrName] = defaultBURef;
                }
            }
            return result;
        }

        private EntityReference GetEntityReference(CrmContext context, XmlElement attrNode, LookupAttributeMetadata attrMeta)
        {
            // Create List of all possible Entity References to return
            List<EntityReference> possibles = new List<EntityReference>();

            // First, see if there is an <ent> element within this <attr> element
            XmlElement entityNode = attrNode.SelectSingleNode("ent") as XmlElement;

            if (entityNode == null)
            {
                // If not, assume Inner Text is Primary Name Value, and reference could be any of the types supported by this attribute, with an unknown GUID
                if (string.IsNullOrWhiteSpace(attrNode.InnerText)) return null;
                possibles.AddRange(attrMeta.Targets.Select(q => new EntityReference(q, Guid.Empty) { Name = attrNode.InnerText }));
            }
            else
            {
                // Otherwise, just add a single "possible"
                EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, entityNode.GetAttribute("ln"));
                XmlElement primaryNode = entityNode.SelectSingleNode("attr[@ln='" + entityMeta.PrimaryNameAttribute + "']") as XmlElement;
                possibles.Add(new EntityReference(entityMeta.LogicalName, entityNode.GetAttribute("id").ToGuid()) { Name = (primaryNode == null) ? null : primaryNode.InnerText });
            }

            // Also add any mappings from IdMappings collection into "possibles"
            possibles.AddRange(possibles.Where(q => this.IdMappings.ContainsKey(q.Id)).Select(q => new EntityReference(q.LogicalName, this.idMappings[q.Id])).ToArray());

            // Try to find a known existing reference. If none of the "possibles" has an ID supplied, allow match by Name, otherwise allow match by ID only at this stage
            EntityReference result = null;
            if (possibles.Any(q => q.Id != Guid.Empty))
            {
                result = knownExist.Where(q => possibles.Any(x => q.Id == x.Id && q.LogicalName == x.LogicalName)).FirstOrDefault();
            }
            else
            {
                result = knownExist.Where(q => possibles.Any(x => q.Name == x.Name && q.LogicalName == x.LogicalName)).FirstOrDefault();
            }
            if (result != null) return result;

            // If a reference was not found, we may need to go to CRM before attempting again

            // Remove any empty references from the list of "possibles"
            IEnumerable<EntityReference> emptyReferences = possibles.Where(q => q.Id == Guid.Empty && string.IsNullOrWhiteSpace(q.Name));
            if (emptyReferences.Any()) possibles = possibles.Except(emptyReferences).ToList();
            if (!possibles.Any()) return null;

            // A query to CRM is required if "possibles" contains any references with no ID, or with IDs not yet known not to exist
            if (possibles.Any(q => q.Id == Guid.Empty || !knownNonExist.Any(x => q.Id == x.Id && q.LogicalName == x.LogicalName)))
            {
                // Look for records in CRM using the IDs and Names supplied in the list of "possibles"
                foreach (string entityType in possibles.Select(q => q.LogicalName).Distinct())
                {
                    EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, entityType);
                    IEnumerable<EntityReference> possiblesOfThisType = possibles.Where(q => q.LogicalName == entityType);
                    IEnumerable<Guid> queryIds = possiblesOfThisType.Select(q => q.Id).Distinct();
                    IEnumerable<string> queryNames = possiblesOfThisType.Where(q => !string.IsNullOrEmpty(q.Name)).Select(q => q.Name).Distinct();

                    QueryExpression query = new QueryExpression(entityType) { ColumnSet = new ColumnSet(entityMeta.PrimaryNameAttribute) };
                    query.Criteria.FilterOperator = LogicalOperator.Or;
                    if (queryIds.Any()) query.Criteria.AddCondition(entityMeta.PrimaryIdAttribute, ConditionOperator.In, queryIds.Cast<object>().ToArray());
                    if (queryNames.Any()) query.Criteria.AddCondition(entityMeta.PrimaryNameAttribute, ConditionOperator.In, queryNames.Cast<object>().ToArray());

                    // Additional filter for Users
                    if (entityMeta.LogicalName == "systemuser")
                    {
                        FilterExpression fe = new FilterExpression();
                        fe.AddCondition("isdisabled", ConditionOperator.Equal, false);
                        fe.AddCondition("accessmode", ConditionOperator.Equal, 0);
                        query.Criteria.AddFilter(fe);
                    }

                    // Retrieve new Entity References
                    IEnumerable<EntityReference> newReferences = context.OrganizationService.RetrieveMultiple(query).Entities.Select(q => new EntityReference(q.LogicalName, q.Id) { Name = q.GetAttributeValue<string>(entityMeta.PrimaryNameAttribute) });

                    // Add to lists of known existing and known non-existing records
                    knownExist.AddRange(newReferences.Where(q => !knownExist.Any(x => q.Id == x.Id && q.LogicalName == x.LogicalName)));
                    knownNonExist.AddRange(queryIds.Where(q => !knownNonExist.Any(x => q == x.Id && x.LogicalName == entityType)).Select(q => new EntityReference(entityType, q)));
                }

                // Try again to find a known existing reference by ID now more records have been retrieved
                if (possibles.Any(q => q.Id != Guid.Empty))
                {
                    result = knownExist.Where(q => possibles.Any(x => q.Id == x.Id && q.LogicalName == x.LogicalName)).FirstOrDefault();
                }
            }

            // Finally, allow match by name even if ID has been supplied, as ID does not exist in CRM
            if (result == null)
            {
                result = knownExist.Where(q => possibles.Any(x => q.Name == x.Name && q.LogicalName == x.LogicalName)).FirstOrDefault();
            }

            // Return
            return result;
        }
    }
}