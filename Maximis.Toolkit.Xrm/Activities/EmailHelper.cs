using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Maximis.Toolkit.Xrm.Activities
{
    public static class EmailHelper
    {
        public static Email ConstructEmailFromTemplate(IOrganizationService orgService, EntityReference regarding, string templateName)
        {
            // Get the Template
            QueryExpression query = new QueryExpression("template") { ColumnSet = new ColumnSet("ownerid") };
            query.Criteria.AddCondition("title", ConditionOperator.Equal, templateName);
            Entity templateEntity = QueryHelper.RetrieveSingleEntity(orgService, query);

            InstantiateTemplateResponse rsp = (InstantiateTemplateResponse)orgService.Execute(new InstantiateTemplateRequest
            {
                ObjectId = regarding.Id,
                ObjectType = regarding.LogicalName,
                TemplateId = templateEntity.Id
            });

            if (rsp.EntityCollection == null || rsp.EntityCollection.Entities == null || rsp.EntityCollection.Entities.Count == 0)
            {
                return null;
            }

            Entity emailEntity = rsp.EntityCollection.Entities[0];
            return new Email
            {
                From = templateEntity.GetAttributeValue<EntityReference>("ownerid"),
                Subject = emailEntity.GetAttributeValue<string>("subject"),
                BodyHtml = emailEntity.GetAttributeValue<string>("description"),
                Regarding = regarding
            };
        }

        public static Entity CreateEmailEntity(IOrganizationService orgService, Email email, bool sendAfterCreate)
        {
            Entity emailEntity = CreateEmailEntityWorker(orgService, email);

            if (sendAfterCreate)
            {
                orgService.Execute(new SendEmailRequest
                {
                    EmailId = emailEntity.Id,
                    IssueSend = true,
                    TrackingToken = email.TrackingToken
                });
            }

            return emailEntity;
        }

        private static Entity CreateEmailEntityWorker(IOrganizationService orgService, Email email)
        {
            // Create Email Entity
            Entity emailEntity = new Entity("email");

            emailEntity["from"] = ActivityHelper.GetActivityPartyArray(ActivityPartyType.Sender, email.From);
            if (email.To != null) emailEntity["to"] = ActivityHelper.GetActivityPartyArray(ActivityPartyType.ToRecipient, email.To.ToArray());
            if (email.Cc != null) emailEntity["cc"] = ActivityHelper.GetActivityPartyArray(ActivityPartyType.CCRecipient, email.Cc.ToArray());
            if (email.Bcc != null) emailEntity["bcc"] = ActivityHelper.GetActivityPartyArray(ActivityPartyType.BccRecipient, email.Bcc.ToArray());

            emailEntity["subject"] = email.Subject;
            emailEntity["description"] = email.BodyHtml;
            emailEntity["regardingobjectid"] = email.Regarding;
            emailEntity["directioncode"] = true;
            emailEntity.Id = orgService.Create(emailEntity);

            // Add any attachments
            if (email.Attachments != null)
            {
                EntityReference emailRef = emailEntity.ToEntityReference();
                foreach (ActivityAttachment attachment in email.Attachments)
                {
                    ActivityHelper.CreateAttachment(orgService, attachment, emailRef);
                }
            }

            return emailEntity;
        }
    }
}