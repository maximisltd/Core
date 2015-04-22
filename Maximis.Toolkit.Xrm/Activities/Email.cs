using Microsoft.Xrm.Sdk;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm.Activities
{
    public class Email
    {
        public List<ActivityAttachment> Attachments { get; set; }

        public List<EntityReference> Bcc { get; set; }

        public string BodyHtml { get; set; }

        public List<EntityReference> Cc { get; set; }

        public EntityReference From { get; set; }

        public EntityReference Regarding { get; set; }

        public string Subject { get; set; }

        public List<EntityReference> To { get; set; }

        public string TrackingToken { get; set; }
    }
}