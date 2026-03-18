using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleParallelProgress : PluginBase
    {
        // Distribution Status OptionSet Values
        private const int NotStarted = 905200000;
        private const int IsPending = 905200001;
        private const int Complete = 905200002;
        private const int Rejected = 905200005;
        private const string DistStatus = "cr8d2_distributionstatus";

        // Routing Status OptionSet Value
        private const int ReviewComplete = 905200002;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Routing Type OptionSet Value
        private const int Parallel = 905200001;
        private const string RoutType = "cr8d2_routingtype";

        // Workflow Status OptionSet Value
        private const int PendingInitiatorAction = 905200012;
        private const string FlowStatus = "cr8d2_workflowstatus";

        // Handle Reject Response
        private const int RejectedByReviewer = 905200006;
        private const int WorkflowTerminated = 905200015;

        // Entity References
        private const string ParentEntityName = "cr8d2_routingsummary";
        private const string ChildEntityName = "cr8d2_documentrouterdecision";

        // Handle Order
        private const string ParentId = "cr8d2_routingsummary";
        private const string SetOrder = "cr8d2_order";

        // Routing summary fields to set actionwith and actionnext
        private const string ActionWith = "cr8d2_actionwith";
        private const string ActionNext = "cr8d2_actionnext";

        // Owner Email
        private const string OwnerEmail = "cr8d2_owneremail";

        // Reviewer Approver lookup fields
        private const string ReviewerLookup = "cr8d2_distributionname";

        public HandleParallelProgress()
            : base(typeof(HandleParallelProgress))
        {
            // Not Implemented
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            tracer.Trace("StartParallelReviewerProgress");

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

                // Check routing type is parallel add get owner email
                Entity parent = sysService.Retrieve(ParentEntityName, parentReference.Id, new ColumnSet(RoutType, OwnerEmail));
                if (!parent.Contains(RoutType) || parent.GetAttributeValue<OptionSetValue>(RoutType).Value != Parallel)

                {
                    tracer.Trace("Routing Type is not Parallel. Exiting.");
                    return;
                }

                // If rejected or completed
                if (postDistributionStatus.Value == Rejected || postDistributionStatus.Value == Complete)
                {
                    tracer.Trace("Reviewer Completed or Rejected. Check for pending reviewers.");

                    // Get remaining active and value exists in list of values notstarted ispending
                    QueryExpression queryremainingReviewers = new QueryExpression(ChildEntityName)
                    {
                        ColumnSet = new ColumnSet(DistStatus, ReviewerLookup),
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

                    EntityCollection remainingReviewers = sysService.RetrieveMultiple(queryremainingReviewers);
                    string ownerEmail = parent.GetAttributeValue<string>(OwnerEmail);

                    if (remainingReviewers.Entities.Count > 0)
                    {
                        tracer.Trace($"{remainingReviewers.Entities.Count} remainingReviewers. Updating ActionWith.");

                        System.Collections.Generic.List<string> pendingNames = new System.Collections.Generic.List<string>();
                        foreach (var rev in remainingReviewers.Entities)
                        {
                            var revRef = rev.GetAttributeValue<EntityReference>(ReviewerLookup);
                            if (revRef != null)
                            {
                                pendingNames.Add(revRef.Name);
                            }
                        }


                        Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentUpdate[ActionWith] = string.Join(", ", pendingNames);
                        parentUpdate[ActionNext] = ownerEmail;

                        sysService.Update(parentUpdate);
                        return;
                    }                
                    else
                    {
                        // No additional reviewers found. Review is complete.  
                        tracer.Trace("No additional reviewers. Review complete");

                        Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentUpdate[RoutStatus] = new OptionSetValue(ReviewComplete);
                        parentUpdate[FlowStatus] = new OptionSetValue(PendingInitiatorAction);
                        parentUpdate[ActionWith] = ownerEmail;
                        parentUpdate[ActionNext] = "Pending";

                        sysService.Update(parentUpdate);
                    }
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Error in HandleParallelProgress: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }




        }
    }
}
