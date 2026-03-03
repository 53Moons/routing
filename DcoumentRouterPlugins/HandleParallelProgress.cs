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

                // If rejected
                if (postDistributionStatus.Value == Rejected)
                {
                    tracer.Trace("Reviewer Rejected. Terminating Workflow.");

                    Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                    parentUpdate[FlowStatus] = new OptionSetValue(WorkflowTerminated);
                    parentUpdate[RoutStatus] = new OptionSetValue(RejectedByReviewer);

                    sysService.Update(parentUpdate);
                    return;
                }

                // If completed
                if (postDistributionStatus.Value == Complete)
                {
                    tracer.Trace("Reviewer Completed. Check for other pending reviewers.");

                    // Keep checking for IsPending or Complete 
                    // note: filter expression and condition expression may not work
                    QueryExpression queryremainingReviewers = new QueryExpression(ChildEntityName)
                    {
                        ColumnSet = new ColumnSet(DistStatus),
                        Criteria = new FilterExpression(LogicalOperator.And)
                        {
                            Conditions =
                            {
                                new ConditionExpression(ParentId, ConditionOperator.Equal, parentReference.Id),
                                new ConditionExpression(DistStatus, ConditionOperator.In, NotStarted, IsPending)
                            }
                        }
                    };                                  

                    EntityCollection remainingReviewers = sysService.RetrieveMultiple(queryremainingReviewers);
                    if (remainingReviewers.Entities.Count > 0)
                    {
                        // Do nothing if there are still reviewers remaining
                        tracer.Trace($"{remainingReviewers.Entities.Count} remainingReviewers");
                        return;
                    }
                    else
                    {
                        // No additional reviewers found. Review is complete.  
                        tracer.Trace("No additional reviewers. Review complete");

                        Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentUpdate[RoutStatus] = new OptionSetValue(ReviewComplete);
                        parentUpdate[FlowStatus] = new OptionSetValue(PendingInitiatorAction);

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
