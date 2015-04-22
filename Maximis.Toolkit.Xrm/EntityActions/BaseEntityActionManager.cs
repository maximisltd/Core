using Maximis.Toolkit.Xrm.Annotations;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm.EntityActions
{
    /// <summary>
    /// Provides a framework for performing multiple discreet actions on an Entity record.
    /// </summary>
    public abstract class BaseEntityActionManager
    {
        /// <summary>
        /// List of actions which are performed against an Entity.
        /// </summary>
        protected List<BaseEntityAction> actions = new List<BaseEntityAction>();

        /// <summary>
        /// Provides access to the CRM Web Service API.
        /// </summary>
        protected IOrganizationService orgService;

        /// <summary>
        /// Allows error messages to be correctly presented on the CRM front-end.
        /// </summary>
        protected ITracingService tracingService;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="orgService">Provides access to the CRM Web Service API.</param>
        /// <param name="tracingService">
        /// Allows error messages to be correctly presented on the CRM front-end.
        /// </param>
        public BaseEntityActionManager(IOrganizationService orgService, ITracingService tracingService)
        {
            this.orgService = orgService;
            this.tracingService = tracingService;
        }

        /// <summary>
        /// Performs all configured actions on the supplied Entity
        /// </summary>
        /// <param name="entity">The entity record on which to perform all actions.</param>
        /// <param name="columns">
        /// Used to ensure that attributes being set were in the initial query
        /// </param>
        /// <returns></returns>
        public void PerformAllActions(Entity entity, ColumnSet columns)
        {
            // Create an Action Definition
            EntityActionDefinition actionDef = new EntityActionDefinition
            {
                OriginalEntity = entity,
                UpdatedEntity = new Entity(entity.LogicalName) { Id = entity.Id },
                OrgService = orgService,
                TracingService = tracingService,
            };

            // Perform each configured action
            foreach (BaseEntityAction action in actions)
            {
                try
                {
                    action.PerformAction(actionDef);
                }
                catch (Exception ex)
                {
                    // If action fails, add a Note to the Entity to record this for debugging purposes
                    AnnotationHelper.CreateAnnotation(orgService, new Annotation
                    {
                        Regarding = entity.ToEntityReference(),
                        Subject = string.Format("Error performing Entity Action '{0}'", action.GetType().Name),
                        NoteText = ex.ToString()
                    });
                }
            }

            // Update if required
            UpdateHelper.SmartUpdate(actionDef.OrgService, actionDef.UpdatedEntity, actionDef.OriginalEntity);
        }
    }
}