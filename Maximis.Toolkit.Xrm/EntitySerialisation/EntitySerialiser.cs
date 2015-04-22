using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Maximis.Toolkit.Xrm.EntitySerialisation
{
    public class EntitySerialiser
    {
        private static EntitySerialiserScope activityPartyScope = new EntitySerialiserScope { EntityType = "activityparty", Columns = new string[] { "partyid", "participationtypemask", "addressused", "ispartydeleted", "resourcespecid" } };
        private readonly List<Guid> processedRecords = new List<Guid>();
        private DisplayStringOptions formats = new DisplayStringOptions { BoolFalse = "false", BoolTrue = "true", DateFormat = "yyyy-MM-ddTHH:mm:ss.fff" };
        private MetadataCache metaCache;
        private Dictionary<string, RelationshipMetadataBase> relationshipMetadata = new Dictionary<string, RelationshipMetadataBase>();

        public EntitySerialiser(MetadataCache metaCache, List<EntitySerialiserScope> scopes)
        {
            this.metaCache = metaCache;
            this.Scopes = scopes;
        }

        public EntitySerialiser(MetadataCache metaCache, EntitySerialiserScope scope)
        {
            this.metaCache = metaCache;
            this.Scopes = new List<EntitySerialiserScope>();
            this.Scopes.Add(scope);
        }

        public DisplayStringOptions Formats { get { return formats; } set { formats = value; } }

        public List<EntitySerialiserScope> Scopes { get; set; }

        public void Reset()
        {
            processedRecords.Clear();
        }

        public string SerialiseEntity(IOrganizationService orgService, EntityReference toSerialise)
        {
            EntitySerialiserScope scope = GetScope(toSerialise);
            return (scope == null) ? null : SerialiseEntityWorker(orgService, orgService.Retrieve(toSerialise.LogicalName, toSerialise.Id, new ColumnSet(scope.Columns)), scope);
        }

        public string SerialiseEntity(IOrganizationService orgService, Entity toSerialise)
        {
            EntitySerialiserScope scope = GetScope(toSerialise.ToEntityReference());
            return (scope == null) ? null : SerialiseEntityWorker(orgService, toSerialise, scope);
        }

        public void SerialiseEntity(IOrganizationService orgService, EntityReference toSerialise, XmlTextWriter xtw)
        {
            EntitySerialiserScope scope = GetScope(toSerialise);
            if (scope != null) SerialiseEntityWorker(orgService, orgService.Retrieve(toSerialise.LogicalName, toSerialise.Id, new ColumnSet(scope.Columns)), scope, xtw);
        }

        public void SerialiseEntity(IOrganizationService orgService, Entity toSerialise, XmlTextWriter xtw)
        {
            EntitySerialiserScope scope = GetScope(toSerialise.ToEntityReference());
            if (scope != null) SerialiseEntityWorker(orgService, toSerialise, scope, xtw);
        }

        private string GetRelatedEntityType(string entityType, RelationshipMetadataBase relMeta)
        {
            switch (relMeta.RelationshipType)
            {
                case RelationshipType.OneToManyRelationship:
                    OneToManyRelationshipMetadata otmMeta = (OneToManyRelationshipMetadata)relMeta;
                    if (otmMeta.ReferencedEntity == entityType) return otmMeta.ReferencingEntity;
                    if (otmMeta.ReferencingEntity == entityType) return otmMeta.ReferencedEntity;
                    break;

                case RelationshipType.ManyToManyRelationship:
                    throw new Exception("Not yet supported!");
            }
            return null;
        }

        private EntitySerialiserScope GetScope(EntityReference toSerialise)
        {
            if (toSerialise == null) return null;

            if (processedRecords.Contains(toSerialise.Id)) return null;

            EntitySerialiserScope scope = Scopes.SingleOrDefault(q => q.EntityType == toSerialise.LogicalName);
            if (scope == null) return null;

            return scope;
        }

        private string SerialiseEntityWorker(IOrganizationService orgService, Entity toSerialise, EntitySerialiserScope scope)
        {
            StringBuilder result = new StringBuilder();

            using (StringWriter sw = new StringWriter(result))
            using (XmlTextWriter xtw = new XmlTextWriter(sw))
            {
                SerialiseEntityWorker(orgService, toSerialise, scope, xtw);
            }

            return result.ToString();
        }

        private void SerialiseEntityWorker(IOrganizationService orgService, Entity toSerialise, EntitySerialiserScope scope, XmlTextWriter xtw)
        {
            processedRecords.Add(toSerialise.Id);

            EntityMetadata toSerialiseMetadata = metaCache.GetEntityMetadata(orgService, toSerialise.LogicalName);

            xtw.WriteStartElement("ent");
            xtw.WriteAttributeString("id", toSerialise.Id.ToString("N"));
            xtw.WriteAttributeString("dn", toSerialiseMetadata.DisplayName.UserLocalizedLabel.Label);
            xtw.WriteAttributeString("ln", toSerialise.LogicalName);

            if (scope.Columns != null && toSerialiseMetadata.Attributes != null)
            {
                foreach (string column in scope.Columns)
                {
                    if (!toSerialise.Contains(column) || toSerialise[column] == null) continue;

                    AttributeMetadata attrMeta =
                       toSerialiseMetadata.Attributes.SingleOrDefault(q => q.LogicalName == column);

                    if (attrMeta == null) continue;

                    xtw.WriteStartElement("attr");
                    if (attrMeta.DisplayName.UserLocalizedLabel != null)
                    {
                        xtw.WriteAttributeString("dn", attrMeta.DisplayName.UserLocalizedLabel.Label);
                    }
                    xtw.WriteAttributeString("ln", attrMeta.LogicalName);
                    xtw.WriteAttributeString("type", attrMeta.AttributeType.ToString());
                    bool writeStringValue = true;

                    switch (attrMeta.AttributeType)
                    {
                        case AttributeTypeCode.Lookup:
                            writeStringValue = false;
                            EntityReference lookupRef = toSerialise.GetAttributeValue<EntityReference>(column);
                            if (lookupRef != null && lookupRef.Id != Guid.Empty)
                            {
                                // Look for Scope
                                EntitySerialiserScope childScope = GetScope(lookupRef);

                                // If scope not found, or we have already serialised this entity, use a
                                // scope which includes only the Primary name
                                if (processedRecords.Contains(lookupRef.Id) || childScope == null)
                                {
                                    childScope = new EntitySerialiserScope { EntityType = lookupRef.LogicalName };
                                    EntityMetadata lookupRefMetadata = metaCache.GetEntityMetadata(orgService, lookupRef.LogicalName);
                                    if (!string.IsNullOrEmpty(lookupRefMetadata.PrimaryNameAttribute)) childScope.Columns = new[] { lookupRefMetadata.PrimaryNameAttribute };
                                }

                                // Serialise lookup Entity
                                ColumnSet colSet = childScope.Columns == null ? new ColumnSet() : new ColumnSet(childScope.Columns);
                                Entity referencedEntity = QueryHelper.TryRetrieve(orgService, lookupRef.LogicalName, lookupRef.Id, colSet);
                                if (referencedEntity != null)
                                {
                                    SerialiseEntityWorker(orgService, referencedEntity, childScope, xtw);
                                }
                            }
                            break;

                        case AttributeTypeCode.PartyList:
                            writeStringValue = false;
                            EntityCollection partyList = toSerialise.GetAttributeValue<EntityCollection>(column);
                            foreach (Entity activityParty in partyList.Entities)
                            {
                                SerialiseEntityWorker(orgService, activityParty, activityPartyScope, xtw);
                            }
                            break;

                        case AttributeTypeCode.Picklist:
                        case AttributeTypeCode.State:
                        case AttributeTypeCode.Status:

                            // For a Picklist, additionally add the value
                            OptionSetValue osv = toSerialise.GetAttributeValue<OptionSetValue>(column);
                            xtw.WriteAttributeString("val", osv.Value.ToString());
                            break;
                    }

                    if (writeStringValue)
                    {
                        xtw.WriteString(MetadataHelper.GetAttributeValueAsDisplayString(orgService, metaCache, toSerialise, attrMeta.LogicalName, this.Formats));
                    }

                    xtw.WriteEndElement();
                }
            }

            if (scope.Relationships != null)
            {
                bool relTagWritten = false;

                foreach (string relationshipName in scope.Relationships)
                {
                    if (!relationshipMetadata.ContainsKey(relationshipName))
                    {
                        relationshipMetadata.Add(relationshipName,
                            MetadataHelper.GetRelationshipMetadata(orgService, relationshipName));
                    }
                    foreach (
                        Entity relatedEntity in
                            QueryHelper.RetrieveRelatedEntities(orgService, toSerialise.ToEntityReference(),
                                relationshipMetadata[relationshipName]).Entities)
                    {
                        if (!processedRecords.Contains(relatedEntity.Id))
                        {
                            if (!relTagWritten)
                            {
                                xtw.WriteStartElement("rel");
                                relTagWritten = true;
                            }
                            SerialiseEntity(orgService, relatedEntity.ToEntityReference(), xtw);
                        }
                    }
                }
                if (relTagWritten) xtw.WriteEndElement();
            }
            xtw.WriteEndElement();
        }
    }
}