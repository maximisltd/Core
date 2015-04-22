using Microsoft.Xrm.Sdk;

namespace Maximis.Toolkit.Xrm.EntityActions
{
    public class EntityActionDefinition
    {
        public IOrganizationService OrgService { get; set; }

        public Entity OriginalEntity { get; set; }

        public ITracingService TracingService { get; set; }

        public Entity UpdatedEntity { get; set; }
    }
}