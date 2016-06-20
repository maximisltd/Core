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
        private CrmContext context;
        private DisplayStringOptions formats = new DisplayStringOptions { BoolFalse = "false", BoolTrue = "true", DateFormat = "yyyy-MM-ddTHH:mm:ss.fff" };
        private Dictionary<string, RelationshipMetadataBase> relationshipMetadata = new Dictionary<string, RelationshipMetadataBase>();

        public EntitySerialiser(CrmContext context, List<EntitySerialiserScope> scopes)
        {
            this.context = context;
            this.Scopes = scopes;
        }

        public EntitySerialiser(CrmContext context, EntitySerialiserScope scope)
        {
            this.context = context;
            this.Scopes = new List<EntitySerialiserScope>();
            this.Scopes.Add(scope);
        }

        public DisplayStringOptions Formats { get { return formats; } set { formats = value; } }

        public List<EntitySerialiserScope> Scopes { get; set; }

        public void Reset()
        {
            processedRecords.Clear();
        }

        public string SerialiseEntity(EntityReference toSerialise)
        {
            EntitySerialiserScope scope = GetScope(toSerialise);
            return (scope == null) ? null : SerialiseEntityWorker(context.OrganizationService.Retrieve(toSerialise.LogicalName, toSerialise.Id, new ColumnSet(scope.Columns)), scope);
        }

        public string SerialiseEntity(Entity toSerialise)
        {
            EntitySerialiserScope scope = GetScope(toSerialise.ToEntityReference());
            return (scope == null) ? null : SerialiseEntityWorker(toSerialise, scope);
        }

        public void SerialiseEntity(EntityReference toSerialise, XmlTextWriter xtw)
        {
            EntitySerialiserScope scope = GetScope(toSerialise);
            if (scope != null) SerialiseEntityWorker(context.OrganizationService.Retrieve(toSerialise.LogicalName, toSerialise.Id, new ColumnSet(scope.Columns)), scope, xtw, false);
        }

        public void SerialiseEntity(Entity toSerialise, XmlTextWriter xtw)
        {
            EntitySerialiserScope scope = GetScope(toSerialise.ToEntityReference());
            if (scope != null) SerialiseEntityWorker(toSerialise, scope, xtw);
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
            EntitySerialiserScope scope = Scopes.SingleOrDefault(q => q.EntityType == toSerialise.LogicalName);
            return scope;
        }

        private string SerialiseEntityWorker(Entity toSerialise, EntitySerialiserScope scope)
        {
            StringBuilder result = new StringBuilder();

            using (StringWriter sw = new StringWriter(result))
            using (XmlTextWriter xtw = new XmlTextWriter(sw))
            {
                SerialiseEntityWorker(toSerialise, scope, xtw);
            }

            return result.ToString();
        }

        private void SerialiseEntityWorker(Entity toSerialise, EntitySerialiserScope scope, XmlTextWriter xtw, bool restrictToScope = false, bool followRelationships = true)
        {
            if (!processedRecords.Contains(toSerialise.Id)) processedRecords.Add(toSerialise.Id);

            EntityMetadata toSerialiseMetadata = MetadataHelper.GetEntityMetadata(context, toSerialise.LogicalName);

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
                                if (restrictToScope || processedRecords.Contains(lookupRef.Id) || childScope == null)
                                {
                                    childScope = new EntitySerialiserScope { EntityType = lookupRef.LogicalName };
                                    EntityMetadata lookupRefMetadata = MetadataHelper.GetEntityMetadata(context, lookupRef.LogicalName);
                                    if (!string.IsNullOrEmpty(lookupRefMetadata.PrimaryNameAttribute)) childScope.Columns = new[] { lookupRefMetadata.PrimaryNameAttribute };
                                }

                                // Serialise lookup Entity
                                ColumnSet colSet = childScope.Columns == null ? new ColumnSet() : new ColumnSet(childScope.Columns);
                                Entity referencedEntity = QueryHelper.TryRetrieve(context.OrganizationService, lookupRef.LogicalName, lookupRef.Id, colSet);
                                if (referencedEntity != null)
                                {
                                    SerialiseEntityWorker(referencedEntity, childScope, xtw, restrictToScope);
                                }
                            }
                            break;

                        case AttributeTypeCode.PartyList:
                            writeStringValue = false;
                            EntityCollection partyList = toSerialise.GetAttributeValue<EntityCollection>(column);
                            foreach (Entity activityParty in partyList.Entities)
                            {
                                SerialiseEntityWorker(activityParty, activityPartyScope, xtw, restrictToScope);
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
                        xtw.WriteString(MetadataHelper.GetAttributeValueAsDisplayString(context, toSerialise, attrMeta.LogicalName, this.Formats));
                    }

                    xtw.WriteEndElement();
                }

                if (!restrictToScope)
                {
                    // Deal with other attributes that are returned in spite of not being in scope
                    // This was written to cope with Calendar Rules but may be useful on other Entities
                    foreach (string attribute in toSerialise.Attributes.Keys.Except(scope.Columns))
                    {
                        EntityCollection embeddedEntities = toSerialise[attribute] as EntityCollection;
                        if (embeddedEntities != null)
                        {
                            xtw.WriteStartElement("attr");
                            xtw.WriteAttributeString("ln", attribute);
                            xtw.WriteAttributeString("type", "EntityCollection");
                            foreach (Entity embeddedEntity in embeddedEntities.Entities)
                            {
                                EntitySerialiserScope embeddedScope = GetScope(embeddedEntity.ToEntityReference());
                                if (embeddedScope != null)
                                {
                                    SerialiseEntityWorker(embeddedEntity, embeddedScope, xtw, true);
                                }
                            }
                            xtw.WriteEndElement();
                        }
                    }
                }
            }

            if (scope.Relationships != null && followRelationships)
            {
                foreach (string relationshipName in scope.Relationships)
                {
                    xtw.WriteStartElement("rel");
                    xtw.WriteAttributeString("name", relationshipName);

                    if (!relationshipMetadata.ContainsKey(relationshipName))
                    {
                        relationshipMetadata.Add(relationshipName,
                            MetadataHelper.GetRelationshipMetadata(context, scope.EntityType, relationshipName));
                    }

                    EntityCollection relatedEntities = QueryHelper.RetrieveRelatedEntities(context.OrganizationService, toSerialise.ToEntityReference(), relationshipMetadata[relationshipName]);
                    foreach (Entity relatedEntity in relatedEntities.Entities)
                    {
                        EntitySerialiserScope relScope = GetScope(relatedEntity.ToEntityReference());
                        if (relScope != null)
                        {
                            Entity relatedEntityFull = context.OrganizationService.Retrieve(relatedEntity.LogicalName, relatedEntity.Id, new ColumnSet(relScope.Columns));
                            SerialiseEntityWorker(relatedEntityFull, relScope, xtw, false, false);
                        }
                    }
                    xtw.WriteEndElement();
                }
            }
            xtw.WriteEndElement();
        }
    }
}