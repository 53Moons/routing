using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleSerialApproverProgressPlugin : PluginBase
    {
        // Approver Reference
        private const string ParentEntityName = "cr8d2_routingsummary";
        private const string ApproverEntityName = "cr8d2_documentroutermanagerdistribution";
        private const string ParentId = "cr8d2_routingsummary";
        private const string SetOrder = "cr8d2_order";

        // Distribution Status OptionSet Values
        private const string DistStatus = "cr8d2_distributionstatus";
        private const int NotStarted = 905200000;
        private const int IsPending = 905200001;
        private const int Complete = 905200002;
        private const int Rejected = 905200005;

        // Routing Status OptionSet Values          
        private const int ReviewComplete = 905200002;
        private const int AllRoutingComplete = 905200004;
        private const int RejectedByApprover = 905200005;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Workflow Status OptionSet Values
        private const int FlowComplete = 905200002;
        private const int PendingInitiatorAction = 905200012;
        private const int Terminated = 905200015;
        private const string FlowStatus = "cr8d2_workflowstatus";

        public HandleSerialApproverProgressPlugin() : base(typeof(HandleSerialApproverProgressPlugin)) { }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            tracer.Trace("Start HandleSerialApproverProgressPlugin");

            if (context.MessageName != "Update" || context.Stage != 40)
                return;
            try
            {

                if (!context.PostEntityImages.TryGetValue("PostImage", out Entity postImage))
                    throw new InvalidPluginExecutionException("PostImage is required.");
                if (!context.PreEntityImages.TryGetValue("PreImage", out Entity preImage))
                    throw new InvalidPluginExecutionException("PreImage is required.");

                // Check dist status pending
                var preDistStatus = preImage.GetAttributeValue<OptionSetValue>(DistStatus);
                var postDistStatus = postImage.GetAttributeValue<OptionSetValue>(DistStatus);

                if (preDistStatus == null || postDistStatus == null || preDistStatus.Value != IsPending)
                {
                    tracer.Trace("Previous Distribution Status was not IsPending. Exiting.");
                    return;
                }

                if (postDistStatus.Value != Complete && postDistStatus.Value != Rejected)
                {
                    tracer.Trace("Status changed, but not to Complete or Rejected. Exiting.");
                    return;
                }

                // Verify routing summary parent reference
                var parentReference = postImage.GetAttributeValue<EntityReference>(ParentId);
                if (parentReference == null)
                    throw new InvalidPluginExecutionException($"Parent routing lookup missing from distribution.");

                // Handle rejection
                if (postDistStatus.Value == Rejected)
                {
                    tracer.Trace("Approver Rejected. Terminating Workflow.");

                    Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                    parentUpdate[FlowStatus] = new OptionSetValue(Terminated);
                    parentUpdate[RoutStatus] = new OptionSetValue(RejectedByApprover);

                    sysService.Update(parentUpdate);
                    return;
                }

                // Find next approver if prev complete
                if (postDistStatus.Value == Complete)
                {
                    tracer.Trace("Approver Completed. Finding next Approver.");

                    QueryExpression queryNextApprover = new QueryExpression(ApproverEntityName)
                    {
                        ColumnSet = new ColumnSet(DistStatus),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                        {
                            new ConditionExpression(ParentId, ConditionOperator.Equal, parentReference.Id),
                            new ConditionExpression(DistStatus, ConditionOperator.Equal, NotStarted),
                            new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active
                        }
                        },
                        TopCount = 1
                    };
                    queryNextApprover.AddOrder(SetOrder, OrderType.Ascending);

                    EntityCollection nextApprovers = sysService.RetrieveMultiple(queryNextApprover);

                    if (nextApprovers.Entities.Count > 0)
                    {
                        // Update next approver to pending
                        Entity nextApproverUpdate = new Entity(ApproverEntityName, nextApprovers.Entities[0].Id);
                        nextApproverUpdate[DistStatus] = new OptionSetValue(IsPending);
                        sysService.Update(nextApproverUpdate);

                        tracer.Trace("Next approver updated to IsPending.");
                    }
                    else
                    {
                        // No more approvers then routing is complete
                        tracer.Trace("No additional approvers. Routing complete.");

                        Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentUpdate[RoutStatus] = new OptionSetValue(AllRoutingComplete);
                        parentUpdate[FlowStatus] = new OptionSetValue(FlowComplete);

                        sysService.Update(parentUpdate);
                    }
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Exception in HandleSerialApproverProgressPlugin: {ex.Message}");
                throw;
            }
        }
    }
}