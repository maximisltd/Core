using Microsoft.Xrm.Sdk;

namespace Maximis.Toolkit.Xrm.EntityActions
{
    public abstract class BaseEntityAction
    {
        public abstract void PerformAction(EntityActionDefinition actionDef);

        protected static int GetOptionSetValue(Entity entity, string attributeName)
        {
            OptionSetValue osVal = entity.GetAttributeValue<OptionSetValue>(attributeName);
            return osVal == null ? 0 : osVal.Value;
        }

        protected bool CompareEntityReferences(EntityReference r1, EntityReference r2)
        {
            if (r1 == null && r2 != null) return false;
            if (r2 == null && r1 != null) return false;
            return (r1.Id == r2.Id);
        }
    }
}