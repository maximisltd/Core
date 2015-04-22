using Microsoft.Xrm.Sdk;
using System;

namespace Maximis.Toolkit.Xrm
{
    public enum EntityImageType
    {
        PreOperation,
        PostOperation
    }

    public enum PluginCollection
    {
        Input,
        Output
    }

    public enum PluginStage
    {
        PreValidation = 10,
        PreOperation = 20,
        MainOperation = 30,
        PostOperation = 40,
        PostOperationCRM4 = 50
    }

    public abstract class BasePlugin : IPlugin
    {
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

        /// <summary>
        /// Wrapper method which extracts the IPluginExecutionContext and ITracingService from the
        /// IServiceProvider, then calls <see cref="ExecutePlugin">ExecutePlugin</see>.
        /// </summary>
        /// <param name="serviceProvider">
        /// Instance of IServiceProvider provided by CRM when the plugin is fired.
        /// </param>
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context =
                (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            ExecutePlugin(serviceProvider, context, tracingService);
        }

        /// <summary>
        /// Raises an error if the plugin is executing in an incorrect stage. <param
        /// name="context">The plugin context.</param><param name="correctStage">The stage in which
        /// the plugin should be executing.</param>
        /// </summary>
        protected void EnsureCorrectPluginStage(IPluginExecutionContext context, PluginStage correctStage)
        {
            PluginStage actualStage = GetPluginStage(context);
            if (actualStage != correctStage)
            {
                throw new InvalidPluginExecutionException(
                    string.Format("Plugin is designed to run in {0} stage, but is running in {1} stage", correctStage,
                        actualStage));
            }
        }

        /// <summary>
        /// Abstract method to contain plugin logic. <param name="serviceProvider">Instance of
        /// IServiceProvider provided by CRM when the plugin is fired.</param><param
        /// name="context">The plugin context.</param><param name="tracingService">Tracing service
        /// allowing error messages to be presented on the CRM front-end.</param>
        /// </summary>
        protected abstract void ExecutePlugin(IServiceProvider serviceProvider, IPluginExecutionContext context,
            ITracingService tracingService);

        /// <summary>
        /// Returns the first attribute value found from multiple Entities. Used to read a value
        /// from either the Target entity or an Image. <param name="attributeName">The name of the
        /// attribute to read.</param><param name="entities">Collection of Entities to read from, in order.</param>
        /// </summary>
        protected T GetAttributeValue<T>(string attributeName, params Entity[] entities)
        {
            if (entities.Length == 0) throw new InvalidPluginExecutionException("Please supply at least one Entity to the GetAttributeValue method.");
            foreach (Entity entity in entities)
            {
                if (entity.HasAttributeWithValue(attributeName))
                {
                    return entity.GetAttributeValue<T>(attributeName);
                }
            }
            return default(T);
        }

        /// <summary>
        /// Returns an Entity Image.
        /// </summary>
        /// <param name="context">The plugin context.</param>
        /// <param name="key">The key of the Image.</param>
        /// <param name="type">Enum depicting whether the Image is Pre-Operation or Post-Operation</param>
        /// <param name="errorIfMissing">
        /// Flag to determine if an error is raised if the Image is not present or not an Entity.
        /// </param>
        /// <returns></returns>
        protected Entity GetEntityImage(IPluginExecutionContext context, string key, EntityImageType type,
            bool errorIfMissing = true)
        {
            Entity result =
                GetObjectFromCollection<Entity, Entity>(
                    type == EntityImageType.PreOperation ? context.PreEntityImages : context.PostEntityImages, key);
            if (errorIfMissing && result == null)
            {
                throw new InvalidPluginExecutionException(string.Format("Failed to retrieve {0} entity image '{1}'",
                    type, key));
            }
            return result;
        }

        /// <summary>
        /// Returns an OrganizationService for use in the Plugin <param
        /// name="serviceProvider">Instance of IServiceProvider provided by CRM when the plugin is
        /// fired.</param><param name="context">The plugin context.</param>
        /// </summary>
        protected IOrganizationService GetOrganizationService(IServiceProvider serviceProvider,
            IPluginExecutionContext context)
        {
            IOrganizationServiceFactory factory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            return service;
        }

        /// <summary>
        /// Returns an EntityCollection from the InputParameters or OutputParameters collection.
        /// <param name="context">The plugin context.</param><param name="errorIfMissing">Flag to
        /// determine if an error is raised if the item is not present or not an Entity.</param>
        /// </summary>
        protected T GetParameter<T>(IPluginExecutionContext context, bool errorIfMissing = true, string collectionKey = "Target", PluginCollection collection = PluginCollection.Input)
        {
            T result = GetObjectFromCollection<T, object>(collection == PluginCollection.Input ? context.InputParameters : context.OutputParameters, collectionKey);
            if (errorIfMissing && result == null)
            {
                throw new InvalidPluginExecutionException(string.Format("Failed to retrieve '{0}' {1} parameter as {2}", collectionKey, collection, typeof(T)));
            }
            return result;
        }

        /// <summary>
        /// Returns the Stage in which the Plugin is executing. <param name="context">The plugin context.</param>
        /// </summary>
        protected PluginStage GetPluginStage(IPluginExecutionContext context)
        {
            return (PluginStage)context.Stage;
        }

        /// <summary>
        /// Writes a message to the Tracing service
        /// </summary>
        /// <param name="tracingService">
        /// Tracing service allowing error messages to be presented on the CRM front-end.
        /// </param>
        /// <param name="format">String format of message.</param>
        /// <param name="args">Arguments for string format.</param>
        protected void Trace(ITracingService tracingService, string format, params object[] args)
        {
            tracingService.Trace(format, args);
        }

        /// <summary>
        /// Returns an object from a DataCollection. <param name="collection">The DataCollection
        /// containing the object.</param><param name="key">The key of the object in the DataCollection.</param>
        /// </summary>
        private T1 GetObjectFromCollection<T1, T2>(DataCollection<string, T2> collection, string key)
        {
            if (collection.Contains(key))
            {
                object o = collection[key];
                if (o is T1) return (T1)o;
            }
            return default(T1);
        }
    }
}