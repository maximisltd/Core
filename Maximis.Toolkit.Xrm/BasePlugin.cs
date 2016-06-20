using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Maximis.Toolkit.Xrm
{
    public enum PluginStage
    {
        PreValidation = 10,
        PreOperation = 20,
        PostOperation = 40,
    }

    public abstract class BasePlugin : IPlugin
    {
        protected class PrimaryNameConfiguration
        {
            private bool allowOverride = true;
            public string NameFormat { get; set; }
            public string[] AttributeNames { get; set; }
            public bool AllowOverride { get { return allowOverride; } set { allowOverride = value; } }
            public DisplayStringOptions DisplayStringOptions { get; set; }
        }

        protected string secureConfig;

        protected string unsecureConfig;

        public BasePlugin()
        {
        }

        public BasePlugin(string unsecureConfig, string secureConfig)
        {
            this.unsecureConfig = unsecureConfig;
            this.secureConfig = secureConfig;
        }

        private static readonly Regex unwantedPrimaryNameEnd = new Regex("[^a-z0-9]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex unwantedPrimaryNameStart = new Regex("^[^a-z0-9]*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public enum ExecutionMode
        {
            Synchronous,
            Asynchronous
        }

        /// <summary>
        /// Set a shared variable which can be accessed in other plugins
        /// </summary>
        protected void SetSharedVariable<T>(PluginContext context, string variableName, T value)
        {
            context.PluginExecutionContext.SharedVariables[variableName] = value;
        }

        /// <summary>
        /// Retrieve a shared variable set in another plugin
        /// </summary>
        protected T GetSharedVariable<T>(PluginContext context, string variableName, IEnumerable<IPluginExecutionContext> ancestorContexts = null)
        {
            IEnumerable<KeyValuePair<string, object>> sharedVariables = ancestorContexts == null ?
                context.PluginExecutionContext.SharedVariables : ancestorContexts.SelectMany(q => q.SharedVariables);
            object result = sharedVariables.FirstOrDefault(q => q.Key == variableName);
            if (result == null) return default(T);
            try { return (T)result; }
            catch { return default(T); }
        }

        /// <summary>
        /// Retrieve a shared variable set in another plugin
        /// </summary>
        protected T GetSharedVariable<T>(PluginContext context, string variableName, bool includeAncestorContexts)
        {
            return GetSharedVariable<T>(context, variableName, includeAncestorContexts ? GetAncestorContexts(context) : null);
        }

        /// <summary>
        /// Returns a collection of ancestor IPluginExecutionContexts.
        /// </summary>
        public List<IPluginExecutionContext> GetAncestorContexts(PluginContext context)
        {
            List<IPluginExecutionContext> result = new List<IPluginExecutionContext>();
            IPluginExecutionContext ancestor = context.PluginExecutionContext.ParentContext;
            while (ancestor != null)
            {
                result.Add(ancestor);
                ancestor = ancestor.ParentContext;
            }
            return result;
        }

        /// <summary>
        /// Wrapper method which creates a PluginContext instance, then calls <see cref="ExecutePlugin">ExecutePlugin</see>.
        /// </summary>
        /// <param name="serviceProvider">
        /// Instance of IServiceProvider provided by CRM when the plugin is fired.
        /// </param>
        public void Execute(IServiceProvider serviceProvider)
        {
            PluginContext context = new PluginContext(serviceProvider, this.GetType());

#if DEBUG
            ExecutePlugin(context);
#else
            try
            {
                ExecutePlugin(context);
            }
            catch (InvalidPluginExecutionException) { throw; }
            catch (Exception ex) { throw new InvalidPluginExecutionException(string.Concat(ex, Environment.NewLine, Environment.NewLine, context.TracingService), ex); }
#endif
        }

        /// <summary>
        /// Abstract method to contain plugin logic.
        /// <param name="context">The plugin context.</param>
        /// </summary>
        protected abstract void ExecutePlugin(PluginContext context);

        /// <summary>
        /// Returns the name value of an Entity Reference. Sometimes this can be null in which case it is retrieved.
        /// <param name="attributeName">The name of the attribute to read.</param>
        /// <param name="entities">Collection of Entities to attempt to read from, in order.</param>
        /// </summary>
        protected string GetEntityReferenceName(PluginContext context, string attributeName, params Entity[] entities)
        {
            // Get Entity Reference
            EntityReference entityRef = GetAttributeValue<EntityReference>(attributeName, entities);

            return GetEntityReferenceName(context, entityRef);
        }

        /// <summary>
        /// Returns the name value of an Entity Reference. Sometimes this can be null in which case it is retrieved.
        /// <param name="attributeName">The name of the attribute to read.</param>
        /// <param name="entities">Collection of Entities to attempt to read from, in order.</param>
        /// </summary>
        protected string GetEntityReferenceName(PluginContext context, EntityReference entityRef)
        {
            // Return null if necessary
            if (entityRef == null) return null;

            // If Name is blank
            if (string.IsNullOrEmpty(entityRef.Name))
            {
                // Get Entity Metadata
                EntityMetadata meta = MetadataHelper.GetEntityMetadata(context, entityRef.LogicalName);

                // Retrieve the Primary Name Attribute of the record and return
                return context.OrganizationService.Retrieve(entityRef.LogicalName, entityRef.Id, new ColumnSet(meta.PrimaryNameAttribute)).GetAttributeValue<string>(meta.PrimaryNameAttribute);
            }

            // Otherwise, return the value
            return entityRef.Name;
        }

        /// <summary>
        /// Returns the first attribute value found from multiple Entities. Used to read a value from either the Target entity or an Image.
        /// <param name="attributeName">The name of the attribute to read.</param>
        /// <param name="entities">Collection of Entities to attempt to read from, in order.</param>
        /// </summary>
        protected T GetAttributeValue<T>(string attributeName, params Entity[] entities)
        {
            if (entities.Length == 0) throw new ArgumentException("GetAttributeValue<T> method called with empty 'entities' argument. Value of 'attributeName': " + attributeName);
            foreach (Entity entity in entities)
            {
                if (entity != null && entity.Contains(attributeName))
                {
                    return entity.GetAttributeValue<T>(attributeName);
                }
            }
            return default(T);
        }

        /// <summary>
        /// Returns an Entity Image.
        /// </summary>
        /// <param name="imageCollection">The collection containing the image.</param>
        /// <param name="key">The key of the Image.</param>
        /// <param name="errorIfMissing"> Flag to determine if an error is raised if the Image is not present.</param>
        protected Entity GetEntityImage(EntityImageCollection imageCollection, string key, bool errorIfMissing = true)
        {
            Entity result = imageCollection.ContainsKey(key) ? imageCollection[key] : null;
            if (errorIfMissing && result == null)
            {
                throw new InvalidPluginExecutionException(string.Format("Failed to retrieve entity image with key '{0}'.", key));
            }
            return result;
        }

        /// <summary>
        /// Returns a Parameter object
        /// </summary>
        ///    <param name="paramCollection">The collection containing the object.</param>
        /// <param name="key">The key of the object.</param>
        /// <param name="errorIfMissing"> Flag to determine if an error is raised if the object is not present.</param>
        protected T GetParameter<T>(ParameterCollection paramCollection, string key, bool errorIfMissing = true)
        {
            object result = paramCollection.ContainsKey(key) ? paramCollection[key] : null;
            Type t = typeof(T);
            if (result == null || result.GetType() != t)
            {
                if (errorIfMissing)
                {
                    string collectionItems = string.Join(", ", paramCollection.Select(q => string.Format("'{0}' ({1})", q.Key, q.Value.GetType().ToString())));
                    throw new InvalidPluginExecutionException(string.Format("Failed to retrieve parameter of type '{0}' with key '{1}'. Available keys and object types are: {2}", t.FullName, key, collectionItems));
                }
                else { return default(T); }
            }
            return (T)result;
        }

        #region Validation

        protected void EnsureNotFutureDate(PluginContext context, string attributeName, Entity target, Entity image = null)
        {
            DateTime val = GetAttributeValue<DateTime>(attributeName, target, image).ToLocalTime();
            if (val > DateTime.Now)
            {
                AttributeMetadata attrMeta = MetadataHelper.GetEntityMetadata(context, target.LogicalName).Attributes.SingleOrDefault(q => q.LogicalName == attributeName);
                string attrDisplayName = attrMeta.LogicalName;
                if (attrMeta.DisplayName.UserLocalizedLabel != null)
                {
                    attrDisplayName = attrMeta.DisplayName.UserLocalizedLabel.Label;
                }
                throw new InvalidPluginExecutionException(string.Format("Date attribute '{0}' cannot be set to a future date.", attrDisplayName));
            }
        }

        /// <summary>
        /// Raises an error if the plugin is executing in an incorrect stage.
        /// <param name="context">The plugin context.</param>
        /// <param name="correctStage">The stage in which the plugin should be executing.</param>
        /// </summary>
        protected string ValidatePluginEntityType(PluginContext context, params string[] entityTypes)
        {
            if (!entityTypes.Contains(context.ExecutionContext.PrimaryEntityName))
            {
                string allowedTypes = string.Join("', '", entityTypes);
                throw new InvalidPluginExecutionException(string.Format("Plugin is designed for these entities: '{0}', but is running against '{1}' entity.", allowedTypes, context.ExecutionContext.PrimaryEntityName));
            }
            return context.ExecutionContext.PrimaryEntityName;
        }

        /// <summary>
        /// Raises an error if the plugin is executing in the incorrect mode.
        /// <param name="context">The plugin context.</param>
        /// <param name="correctMessages">The mode in which the plugin should be executing.</param>
        /// </summary>
        protected ExecutionMode ValidatePluginExecutionMode(PluginContext context, params ExecutionMode[] modes)
        {
            ExecutionMode actualMode = context.ExecutionContext.OperationId == Guid.Empty ? ExecutionMode.Synchronous : ExecutionMode.Asynchronous;
            if (!modes.Contains(actualMode))
            {
                throw new InvalidPluginExecutionException(string.Format("Plugin is designed for execution in {0} mode.", actualMode));
            }
            return actualMode;
        }

        /// <summary>
        /// Raises an error if the plugin is executing in an incorrect stage.
        /// <param name="context">The plugin context.</param>
        /// <param name="correctMessages">The messages for which the plugin should be executing.</param>
        /// </summary>
        protected string ValidatePluginMessage(PluginContext context, params string[] messages)
        {
            if (!messages.Contains(context.ExecutionContext.MessageName))
            {
                string allowedMessages = string.Join("', '", messages);
                throw new InvalidPluginExecutionException(string.Format("Plugin is designed for these messages: '{0}', but is running against '{1}' message.", allowedMessages, context.ExecutionContext.MessageName));
            }
            return context.ExecutionContext.MessageName;
        }

        /// <summary>
        /// Raises an error if the plugin is executing in an incorrect stage.
        /// <param name="context">The plugin context.</param>
        /// <param name="correctStage">The stage in which the plugin should be executing.</param>
        /// </summary>
        protected PluginStage ValidatePluginStage(PluginContext context, params PluginStage[] stages)
        {
            PluginStage actualStage = (PluginStage)context.PluginExecutionContext.Stage;
            if (!stages.Contains(actualStage))
            {
                string allowedStages = string.Join("', '", stages.Select(q => q.ToString()));
                throw new InvalidPluginExecutionException(string.Format("Plugin is designed for these stages: '{0}', but is running against '{1}' stage.", allowedStages, actualStage));
            }
            return actualStage;
        }

        #endregion Validation

        #region FetchXML Value Extraction

        protected DateTime GetDateTimeConditionFromFetchXml(XmlDocument fetchDoc, string attribute, DateTime defaultValue)
        {
            // Set result to default value
            DateTime result = defaultValue;

            // Look for attribute value
            string stringVal = GetStringConditionFromFetchXml(fetchDoc, attribute);

            // if not found, look for alternate attribute value
            if (string.IsNullOrEmpty(stringVal)) stringVal = GetStringConditionFromFetchXml(fetchDoc, attribute + "string");

            // Try to convert to DateTime
            if (!string.IsNullOrEmpty(stringVal)) DateTime.TryParse(stringVal, out result);
            return result;
        }

        protected bool GetBoolConditionFromFetchXml(XmlDocument fetchDoc, string attribute, bool defaultValue)
        {
            // Set result to default value
            bool result = defaultValue;

            // Look for attribute value
            string stringVal = GetStringConditionFromFetchXml(fetchDoc, attribute);

            // Try to convert to Guid
            if (!string.IsNullOrEmpty(stringVal)) bool.TryParse(stringVal, out result);
            return result;
        }

        protected Guid GetGuidConditionFromFetchXml(XmlDocument fetchDoc, string attribute)
        {
            // Set result to default value
            Guid result = Guid.Empty;

            // Look for attribute value
            string stringVal = GetStringConditionFromFetchXml(fetchDoc, attribute);

            // Try to convert to Guid
            if (!string.IsNullOrEmpty(stringVal)) Guid.TryParse(stringVal, out result);
            return result;
        }

        protected int GetIntConditionFromFetchXml(XmlDocument fetchDoc, string attribute, int defaultValue)
        {
            // Set result to default value
            int result = defaultValue;

            // Look for attribute value
            string stringVal = GetStringConditionFromFetchXml(fetchDoc, attribute);

            // Try to convert to Integer
            if (!string.IsNullOrEmpty(stringVal)) int.TryParse(stringVal, out result);
            return result;
        }

        protected string GetStringConditionFromFetchXml(XmlDocument fetchDoc, string attribute)
        {
            // Get XML Node
            XmlNode result = fetchDoc.SelectSingleNode(string.Format("//condition[@attribute='{0}']", attribute));

            // Return null if not found, or value of "value" attribute
            if (result == null) return null;
            return ((XmlElement)result).GetAttribute("value");
        }

        #endregion FetchXML Value Extraction

        #region Primary Name

        /// <summary>
        /// Sets the Name of an Entity to a value constructed from other attribute values
        /// </summary>
        protected void SetPrimaryNameValue(PluginContext context, PrimaryNameConfiguration config, Entity target, Entity image = null)
        {
            // Get Entity Metadata
            context.TracingService.Trace("Retrieving Entity Metadata.");
            EntityMetadata entityMeta = MetadataHelper.GetEntityMetadata(context, target.LogicalName);

            // If Name is being set in Target, drop out
            if (config.AllowOverride)
            {
                string targetName = target.GetAttributeValue<string>(entityMeta.PrimaryNameAttribute);
                if (!string.IsNullOrWhiteSpace(targetName))
                {
                    context.TracingService.Trace("Allowing explicitly set name - exiting.");
                    return;
                }
            }

            // Throw Exception if Image is present but doesn't include Name
            if (image != null)
            {
                string imageName = image.GetAttributeValue<string>(entityMeta.PrimaryNameAttribute);
                if (string.IsNullOrWhiteSpace(imageName))
                {
                    context.TracingService.Trace("Image primary name attribute is missing or null - checking record.");
                    string actualNameValue = context.OrganizationService.Retrieve(image.LogicalName, image.Id, new ColumnSet(entityMeta.PrimaryNameAttribute)).GetAttributeValue<string>(entityMeta.PrimaryNameAttribute);
                    if (!string.IsNullOrEmpty(actualNameValue))
                    {
                        // If record currently has a value for its primary name attribute, that value should be present in the image. Throw an Exception to warn of this.
                        throw new InvalidPluginExecutionException(string.Format("{0} SetPrimaryNameValue: Image of '{1}' does not include Primary Name Attribute '{2}'", this.GetType().Name, entityMeta.LogicalName, entityMeta.PrimaryNameAttribute));
                    }
                }
            }

            // Get current Name
            string currentName = GetAttributeValue<string>(entityMeta.PrimaryNameAttribute, target, image);

            // Get the Attribute Values needed to construct the new Name
            context.TracingService.Trace("Extracting Attribute Values needed for Name.");
            List<object> attributeValues = new List<object>();
            bool canCreateName = false;
            foreach (string attributeName in config.AttributeNames)
            {
                AttributeMetadata attributeMeta = entityMeta.Attributes.SingleOrDefault(q => q.LogicalName == attributeName);
                if (attributeMeta == null)
                {
                    // Unknown attribute, so insert name in square brackets as a warning
                    attributeValues.Add(string.Format("[{0}]", attributeName));
                    canCreateName = true;
                }
                else
                {
                    // Insert "display string" of value
                    string attributeValue = MetadataHelper.GetAttributeValueAsDisplayString(context, GetAttributeValue<object>(attributeName, target, image), attributeMeta, config.DisplayStringOptions);
                    canCreateName |= !string.IsNullOrWhiteSpace(attributeValue);
                    attributeValues.Add(attributeValue);
                }
            }

            // If Name can be created, create it
            if (canCreateName)
            {
                context.TracingService.Trace("Setting new Name value.");

                // Generate complete name
                string newName = string.Format(config.NameFormat, attributeValues.ToArray()).Trim();

                // If name starts or ends with punctuation, strip out
                // For example if format is "{0} - {1}" or "{0}: {1}" and either value is empty
                newName = unwantedPrimaryNameStart.Replace(newName, string.Empty);
                newName = unwantedPrimaryNameEnd.Replace(newName, string.Empty);

                // Set Primary Name attribute, ensuring length does not exceed allowed
                StringAttributeMetadata primaryMeta = entityMeta.Attributes.SingleOrDefault(q => q.LogicalName == entityMeta.PrimaryNameAttribute) as StringAttributeMetadata;
                newName = newName.TruncateWithEllipsis(primaryMeta.MaxLength.Value);
                if (newName != currentName) target[entityMeta.PrimaryNameAttribute] = newName;
            }
            else
            {
                // Otherwise, set default name, if current name is blank
                if (string.IsNullOrWhiteSpace(currentName))
                {
                    context.TracingService.Trace("Setting default name.");
                    string entityDisplayName = null;
                    if (entityMeta.DisplayName.UserLocalizedLabel != null)
                    {
                        entityDisplayName = entityMeta.DisplayName.UserLocalizedLabel.Label;
                    }
                    if (string.IsNullOrEmpty(entityDisplayName)) entityDisplayName = entityMeta.LogicalName;
                    target[entityMeta.PrimaryNameAttribute] = entityDisplayName;
                }
            }
        }

        #endregion Primary Name
    }
}