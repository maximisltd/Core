using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;

namespace Maximis.Toolkit.Xrm.Annotations
{
    public static class AnnotationHelper
    {
        private static readonly ColumnSet COLSET = new ColumnSet("objectid", "subject", "notetext", "mimetype",
            "filename", "documentbody");

        public static Entity CreateAnnotation(IOrganizationService orgService, Annotation annotation)
        {
            Entity result = new Entity("annotation");
            result["objectid"] = annotation.Regarding;
            result["objecttypecode"] = annotation.Regarding.LogicalName;
            result["subject"] = annotation.Subject;
            result["notetext"] = annotation.NoteText;
            if (annotation.Attachment != null)
            {
                result["mimetype"] = annotation.Attachment.MimeType;
                result["filename"] = annotation.Attachment.Filename;
                if (!string.IsNullOrEmpty(annotation.Attachment.BodyText))
                {
                    result["documentbody"] =
                        Convert.ToBase64String(Encoding.UTF8.GetBytes(annotation.Attachment.BodyText));
                }
                else if (annotation.Attachment.Body != null)
                {
                    result["documentbody"] =
                        Convert.ToBase64String(annotation.Attachment.Body);
                }
            }
            return UpdateHelper.CreateOrUpdate(orgService, result);
        }

        public static Annotation RetrieveAnnotation(IOrganizationService orgService, QueryExpression query)
        {
            query.ColumnSet = COLSET;
            Entity entity = QueryHelper.RetrieveSingleEntity(orgService, query);
            return GetAnnotationObjectFromEntity(entity);
        }

        public static List<Annotation> RetrieveAnnotations(IOrganizationService orgService, QueryExpression query)
        {
            List<Annotation> result = new List<Annotation>();
            query.ColumnSet = COLSET;
            foreach (Entity entity in orgService.RetrieveMultiple(query).Entities)
            {
                result.Add(GetAnnotationObjectFromEntity(entity));
            }
            return result;
        }

        public static List<Annotation> RetrieveAnnotations(IOrganizationService orgService, EntityReference regarding)
        {
            QueryExpression query = new QueryExpression("annotation") { ColumnSet = COLSET };
            query.Criteria.AddCondition("objectid", ConditionOperator.Equal, regarding.Id);
            List<Annotation> result = new List<Annotation>();
            foreach (Entity entity in orgService.RetrieveMultiple(query).Entities)
            {
                result.Add(GetAnnotationObjectFromEntity(entity));
            }
            return result;
        }

        private static Annotation GetAnnotationObjectFromEntity(Entity entity)
        {
            if (entity == null) return null;

            Annotation annotation = new Annotation
            {
                Subject = entity.GetAttributeValue<string>("subject"),
                NoteText = entity.GetAttributeValue<string>("notetext"),
                Regarding = entity.GetAttributeValue<EntityReference>("objectid"),
            };
            if (entity.HasAttributeWithValue("documentbody"))
            {
                byte[] body = Convert.FromBase64String(entity.GetAttributeValue<string>("documentbody"));
                annotation.Attachment = new AnnotationAttachment
                {
                    MimeType = entity.GetAttributeValue<string>("mimetype"),
                    Filename = entity.GetAttributeValue<string>("filename"),
                    BodyText = Encoding.UTF8.GetString(body),
                    Body = body
                };
            }
            return annotation;
        }
    }
}