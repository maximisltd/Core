using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm
{
    public abstract class BaseCustomWorkflowActivity : CodeActivity
    {
        protected static readonly ColumnSet asyncOpCols = new ColumnSet("message", "friendlymessage");
        protected static readonly Guid dummyGuid = Guid.NewGuid();

        public abstract void ExecuteWorkflowStep(WorkflowStepContext context);

        /// <summary>
        /// Returns a collection of ancestor IWorkflowContexts.
        /// </summary>
        public List<IWorkflowContext> GetAncestorContexts(WorkflowStepContext context)
        {
            List<IWorkflowContext> result = new List<IWorkflowContext>();
            IWorkflowContext ancestor = context.WorkflowContext.ParentContext;
            while (ancestor != null) { result.Add(ancestor); ancestor = ancestor.ParentContext; }
            return result;
        }

        protected override void Execute(CodeActivityContext activityContext)
        {
            // Create Plugin Context
            WorkflowStepContext context = new WorkflowStepContext(activityContext, this.GetType());

            // Call ExecuteWorkflowStep method
            ExecuteWorkflowStep(context);
        }
    }
}