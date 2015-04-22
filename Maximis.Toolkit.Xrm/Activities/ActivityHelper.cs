using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Text;

namespace Maximis.Toolkit.Xrm.Activities
{
    public enum ActivityPartyType
    {
        Sender = 1,
        ToRecipient = 2,
        CCRecipient = 3,
        BccRecipient = 4,
        RequiredAttendee = 5,
        OptionalAttendee = 6,
        Organizer = 7,
        Regarding = 8,
        Owner = 9,
        Resource = 10,
        Customer = 11,
        Partner = 12
    }

    public static class ActivityHelper
    {
        public static Entity CreateAttachment(IOrganizationService orgService, ActivityAttachment attachment, EntityReference attachTo)
        {
            Entity result = new Entity("activitymimeattachment");
            result["objectid"] = attachTo;
            result["objecttypecode"] = attachTo.LogicalName;
            result["subject"] = attachment.Title;
            result["mimetype"] = attachment.MimeType;
            result["filename"] = attachment.Filename;

            if (string.IsNullOrEmpty(attachment.BodyText))
            {
                result["body"] = Convert.ToBase64String(attachment.Body);
            }
            else
            {
                result["body"] = Convert.ToBase64String(Encoding.UTF8.GetBytes(attachment.BodyText));
            }

            result.Id = orgService.Create(result);
            return result;
        }

        public static Entity GetActivityParty(ActivityPartyType partyType, EntityReference entityRef)
        {
            if (entityRef == null) return null;

            Entity result = new Entity("activityparty");
            result["partyid"] = entityRef;
            result["participationtypemask"] = new OptionSetValue((int)partyType);
            return result;
        }

        public static Entity[] GetActivityPartyArray(ActivityPartyType partyType, params EntityReference[] entityRefs)
        {
            if (entityRefs == null || entityRefs.Length == 0) return null;

            List<Entity> result = new List<Entity>();
            foreach (EntityReference entityRef in entityRefs)
            {
                Entity partyMember = GetActivityParty(partyType, entityRef);
                if (partyMember != null) result.Add(partyMember);
            }
            return (result.Count > 0) ? result.ToArray() : null;
        }
    }
}