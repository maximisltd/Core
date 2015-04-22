using Microsoft.Xrm.Sdk;
using System.Linq;

namespace Maximis.Toolkit.Xrm.EntitySerialisation
{
    public class EntitySerialiserScope
    {
        public string[] Columns { get; set; }

        public string EntityType { get; set; }

        public string[] Relationships { get; set; }

        public static EntitySerialiserScope CreateFromEntity(Entity entity)
        {
            return new EntitySerialiserScope { EntityType = entity.LogicalName, Columns = entity.Attributes.Select(q => q.Key).ToArray() };
        }
    }
}