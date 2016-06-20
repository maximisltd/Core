using Maximis.Toolkit.Caching;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Maximis.Toolkit.Xrm
{
    public class CrmContext
    {
        private EntityReference defaultCurrency;

        public CrmContext()
        {
        }

        public CrmContext(IOrganizationService orgService)
        {
            this.TracingService = new TracingService(new DiagnosticTracingService());
            this.CacheManager = new CacheManager(Guid.NewGuid().ToString());
            this.OrganizationService = orgService;
        }

        public CrmContext(IOrganizationService orgService, string uniqueName)
        {
            this.TracingService = new TracingService(new DiagnosticTracingService());
            this.CacheManager = new CacheManager(uniqueName);
            this.OrganizationService = orgService;
        }

        public CacheManager CacheManager { get; set; }

        public EntityReference DefaultCurrency
        {
            get
            {
                if (defaultCurrency == null)
                {
                    defaultCurrency = OrganizationService.Retrieve("organization", this.ExecutionContext.OrganizationId,
                        new ColumnSet("basecurrencyid")).GetAttributeValue<EntityReference>("basecurrencyid");
                }
                return defaultCurrency;
            }
        }

        public IExecutionContext ExecutionContext { get; set; }

        public IOrganizationService OrganizationService { get; set; }

        public ITracingService TracingService { get; set; }

        public string GetTraceLog()
        {
            TracingService tracingService = this.TracingService as TracingService;
            if (tracingService == null) return null;
            return tracingService.ToString();
        }

        private class DiagnosticTracingService : ITracingService
        {
            public void Trace(string format, params object[] args)
            {
                System.Diagnostics.Trace.WriteLine(string.Format(format, args));
            }
        }
    }
}