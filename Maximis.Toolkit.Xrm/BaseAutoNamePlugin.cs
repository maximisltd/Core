using Maximis.Toolkit.Html;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm
{
    /// <summary>
    /// Automatically sets the name of an Entity to a value derived from other field values.
    /// </summary>
    public abstract class BaseAutoNamePlugin : BasePlugin
    {
        public string[] AttributeNames { get; set; }

        public string DateFormat { get; set; }

        public string Format { get; set; }

        public int MaxLength { get; set; }

        public string NameAttribute { get; set; }

        /// <summary>
        /// Virtual method allowing a class inheriting from this one to apply custom logic before
        /// the Auto Name operation.
        /// </summary>
        /// <param name="targetEntity">The Entity about to be updated.</param>
        /// <param name="entityImage">
        /// An Entity Image containing the fields used for creating the Name field value.
        /// </param>
        /// <param name="metadata">The Metadata for the Entity Type.</param>
        protected virtual void BeforeNameChange(IOrganizationService orgService, EntityMetadata metadata, Entity targetEntity, Entity entityImage)
        {
        }

        /// <summary>
        /// Executes the plugin logic. <param name="serviceProvider">Instance of IServiceProvider
        /// provided by CRM when the plugin is fired.</param><param name="context">The plugin
        /// context.</param><param name="tracingService">Tracing service allowing error messages to
        /// be presented on the CRM front-end.</param>
        /// </summary>
        protected override void ExecutePlugin(IServiceProvider serviceProvider, IPluginExecutionContext context,
            ITracingService tracingService)
        {
            // Drop out if required information is not present
            if (string.IsNullOrEmpty(NameAttribute)) return;
            if (string.IsNullOrEmpty(Format)) return;
            if (AttributeNames == null || AttributeNames.Length == 0) return;

            // Ensure Plugin is running before operation
            EnsureCorrectPluginStage(context, PluginStage.PreOperation);

            // Get Target Entity
            Entity targetEntity = GetParameter<Entity>(context, false);

            // If it's null, we don't want to set the name (it's probably being deleted), so drop out
            if (targetEntity == null) return;

            // Get Entity Image (if it exists - optional)
            Entity entityImage = GetEntityImage(context, "AutoName", EntityImageType.PreOperation, false);

            // Get the Entity Metadata
            IOrganizationService orgService = GetOrganizationService(serviceProvider, context);
            MetadataCache metaCache = new MetadataCache();
            EntityMetadata metadata = metaCache.GetEntityMetadata(orgService, targetEntity.LogicalName);

            // Optionally call any custom code before the name change happens
            BeforeNameChange(orgService, metadata, targetEntity, entityImage);

            // Get the Field Values we're interested in
            List<object> attributeValues = new List<object>();
            foreach (string attributeName in AttributeNames)
            {
                AttributeMetadata attributeMeta = MetadataHelper.GetAttributeMetadata(metadata, attributeName);
                if (attributeMeta == null)
                {
                    // Unknown attribute, so insert name in square brackets as a warning
                    attributeValues.Add(string.Format("[{0}]", attributeName));
                }
                else
                {
                    // Insert "display string" of value
                    attributeValues.Add(MetadataHelper.GetAttributeValueAsDisplayString(orgService, metaCache, (entityImage == null ? targetEntity : entityImage), attributeMeta.LogicalName, new DisplayStringOptions { DateFormat = this.DateFormat }));
                }
            }

            // Set the Name
            string newName = HtmlHelper.StripHtml(string.Format(Format, attributeValues.ToArray())).TruncateWithEllipsis(this.MaxLength);

            string currentName = targetEntity.GetAttributeValue<string>(this.NameAttribute);
            if (currentName != newName) targetEntity[this.NameAttribute] = newName;
        }
    }
}