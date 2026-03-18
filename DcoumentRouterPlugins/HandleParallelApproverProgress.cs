using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Diagnostics.Tracing;

namespace DcoumentRouterPlugins
{
    public class HandleParallelApproverProgress : PluginBase
    {
        // Distribution Status OptionSet Values
        private const int NotStarted = 905200000;
        private const int IsPending = 905200001;
        private const int Complete = 905200002;
        private const int Rejected = 905200005;
        private const string DistStatus = "cr8d2_distributionstatus";

        // Routing Status OptionSet Value
        private const int RoutingComplete = 905200004;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Routing Type OptionSet Value
        private const int Parallel = 905200001;
        private const string RoutType = "cr8d2_routingtype";

        // Workflow Status OptionSet Value
        private const int PendingInitiatorAction = 905200012;
        private const int FlowComplete = 905200002;
        private const string FlowStatus = "cr8d2_workflowstatus";

        // Handle Reject Response
        private const int RejectedByApprover = 905200005;
        private const int WorkflowTerminated = 905200015;

        // Entity References
        private const string ParentEntityName = "cr8d2_routingsummary";
        private const string ChildEntityName = "cr8d2_documentroutermanagerdistribution";

        // Handle Order
        private const string ParentId = "cr8d2_routingsummary";
        private const string SetOrder = "cr8d2_order";

        // Routing summary fields to set actionwith and actionnext
        private const string ActionWith = "cr8d2_actionwith";
        private const string ActionNext = "cr8d2_actionnext";

        // Approver lookup field
        private const string ApproverLookup = "cr8d2_managername";

        // Owner Email
        private const string OwnerEmail = "cr8d2_owneremail";

        public HandleParallelApproverProgress()
            : base(typeof(HandleParallelApproverProgress))
        {
            // Not Implemented
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            tracer.Trace("StartParallelApproverProgress");

            // Check stage is post operation 40
            if (context.MessageName != "Update" || context.Stage != 40)
                return;

            try
            {

                if (!context.PostEntityImages.TryGetValue("Image", out Entity postImage))
                    throw new Exception("Post Image is required.");
                if (!context.PreEntityImages.TryGetValue("Image", out Entity preImage))
                    throw new Exception("Pre Image is required.");

                // Confirm distribution status in image
                if (!postImage.TryGetAttributeValue(DistStatus, out OptionSetValue postDistributionStatus))
                    throw new Exception("Distribution Status not in Post Image");
                if (!preImage.TryGetAttributeValue(DistStatus, out OptionSetValue preDistributionStatus))
                    throw new Exception("Distribution Status not in Pre Image");

                // Distribution status has to be pending
                if (preDistributionStatus.Value != IsPending)
                {
                    tracer.Trace("Previous Distribution Status was not IsPending. Exiting.");
                    return;
                }

                // Verify completed or rejected
                if (postDistributionStatus.Value != Complete && postDistributionStatus.Value != Rejected)
                {
                    tracer.Trace($"Distribution status changed to {postDistributionStatus.Value}, which is neither Complete nor Rejected. Exiting.");
                    return;
                }

                // Get parent 
                var parentReference = postImage.GetAttributeValue<EntityReference>(ParentId);
                if (parentReference == null)
                {
                    throw new Exception($"Parent routing summary lookup ({ParentId}) missing from distribution.");
                }

                // Check routing type is parallel
                Entity parent = sysService.Retrieve(ParentEntityName, parentReference.Id, new ColumnSet(RoutType, OwnerEmail));
                if (!parent.Contains(RoutType) || parent.GetAttributeValue<OptionSetValue>(RoutType).Value != Parallel)

                {
                    tracer.Trace("Routing Type is not Parallel. Exiting.");
                    return;
                }

                // If rejected
                if (postDistributionStatus.Value == Rejected)
                {
                    tracer.Trace("Approver Rejected. Terminating Workflow.");

                    Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                    parentUpdate[FlowStatus] = new OptionSetValue(WorkflowTerminated);
                    parentUpdate[RoutStatus] = new OptionSetValue(RejectedByApprover);
                    parentUpdate[ActionWith] = "None";
                    parentUpdate[ActionNext] = "None";

                    sysService.Update(parentUpdate);
                    return;
                }

                // If completed
                if (postDistributionStatus.Value == Complete)
                {
                    tracer.Trace("Approver Completed. Check for other pending approvers.");

                    // Keep checking for IsPending or Complete 
                    // note: filter expression and condition expression may not work
                    QueryExpression queryremainingApprovers = new QueryExpression(ChildEntityName)
                    {
                        ColumnSet = new ColumnSet(DistStatus, ApproverLookup),
                        Criteria = new FilterExpression(LogicalOperator.And)
                        {
                            Conditions =
                            {
                                new ConditionExpression(ParentId, ConditionOperator.Equal, parentReference.Id),
                                new ConditionExpression(DistStatus, ConditionOperator.In, NotStarted, IsPending),
                                new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                            }
                        }
                    };

                    EntityCollection remainingApprovers = sysService.RetrieveMultiple(queryremainingApprovers);
                    string ownerEmail = parent.GetAttributeValue<string>(OwnerEmail);
                    
                    if (remainingApprovers.Entities.Count > 0)
                    {
                        tracer.Trace($"{remainingApprovers.Entities.Count} remainingApprovers. Updating ");
                        
                        System.Collections.Generic.List<string> pendingNames = new System.Collections.Generic.List<string>();
                        foreach (var app in remainingApprovers.Entities)
                        {
                            var appRef = app.GetAttributeValue<EntityReference>(ApproverLookup);
                            if (appRef != null)
                            {
                                pendingNames.Add(appRef.Name);
                            }

                        }

                        Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentUpdate[ActionWith] = string.Join(",", pendingNames);
                        parentUpdate[ActionNext] = ownerEmail;

                        sysService.Update(parentUpdate);
                        return;                                               
                      
                    }
                    else
                    {
                        // No additional approvers found. Routing is complete.  
                        tracer.Trace("No additional approvers. Routing complete");

                        Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentUpdate[RoutStatus] = new OptionSetValue(RoutingComplete);
                        parentUpdate[FlowStatus] = new OptionSetValue(FlowComplete);
                        parentUpdate[ActionWith] = ownerEmail;
                        parentUpdate[ActionNext] = "None";

                        sysService.Update(parentUpdate);
                    }
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Error in HandleParallelApproverProgress: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }




        }
    }
}
