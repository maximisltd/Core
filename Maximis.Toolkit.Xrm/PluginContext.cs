using Maximis.Toolkit.Caching;
using Microsoft.Xrm.Sdk;
using System;

namespace Maximis.Toolkit.Xrm
{
    public class PluginContext : CrmContext
    {
        public PluginContext(IServiceProvider serviceProvider, Type type)
        {
            this.PluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this.TracingService = new TracingService((ITracingService)serviceProvider.GetService(typeof(ITracingService)));
            this.CacheManager = new CacheManager(this.PluginExecutionContext.OrganizationName);
            this.OrganizationService = ((IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory))).CreateOrganizationService(this.PluginExecutionContext.UserId);
            this.ExecutionContext = this.PluginExecutionContext;
            this.TracingService.Trace("Plugin Instantiated: {0}", type.Name);
            this.TracingService.Trace("Primary Record and Message: {0} {1} {2}", this.PluginExecutionContext.PrimaryEntityName, this.PluginExecutionContext.PrimaryEntityId, this.PluginExecutionContext.MessageName);
            this.TracingService.Trace("Depth: {0}", this.PluginExecutionContext.Depth);
        }

        public IPluginExecutionContext PluginExecutionContext { get; set; }
    }
}