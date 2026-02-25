using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleSerialReviewerProgressPlugin : PluginBase
    {
        // Distribution Status OptionSet Values
        private const int NotStarted = 905200000;
        private const int IsPending = 905200001;
        private const int Complete = 905200002;
        private const int Rejected = 905200005;
        private const string DistStatus = "cr8d2_distributionstatus";

        // Routing Status OptionSet Value
        private const int ReviewComplete = 905000002;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Workflow Status OptionSet Value
        private const int PendingInitiatorAction = 905200012;
        private const int SerialReviewPending = 905200003;
        private const string FlowStatus = "cr8d2_workflowstatus";

        // Handle Reject Response
        private const int RejectedByReviewer = 905200006;
        private const int WorkflowTerminated = 905200015;

        // Entity References
        private const string ParentEntityName = "cr8d2_routingsummary";
        private const string ChildEntityName = "cr8d2_documentrouterdecision";
        private const string ReviewerLookup = "cr8d2_distributionname";

        // Handle Order
        private const string ParentId = "cr8d2_routingsummaryid";
        private const string SetOrder = "cr8d2_order";

        public HandleSerialReviewerProgressPlugin()
            : base(typeof(HandleSerialReviewerProgressPlugin))
        {
            // Not Implemented
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            tracer.Trace("StartSerialReviewerProgress");

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
                if (
                    !postImage.TryGetAttributeValue(
                        DistStatus,
                        out OptionSetValue postDistributionStatus
                    )
                )
                    throw new Exception("Distribution Status not in Post Image");
                if (
                    !preImage.TryGetAttributeValue(
                        DistStatus,
                        out OptionSetValue preDistributionStatus
                    )
                )
                    throw new Exception("Distribution Status not in Pre Image");

                // Check Workflow Status serial review pending (triggered from initial flow)
                OptionSetValue preWorkflowStatus = preImage.GetAttributeValue<OptionSetValue>(
                    FlowStatus
                );
                OptionSetValue postWorkflowStatus = postImage.GetAttributeValue<OptionSetValue>(
                    FlowStatus
                );

                //bool isSerialReview = (preWorkflowStatus != null && preWorkflowStatus.Value == SerialReviewPending) ||
                //                      (postWorkflowStatus != null && postWorkflowStatus.Value == SerialReviewPending);

                //if (!isSerialReview)
                //{
                //    tracer.Trace("Workflow status is not SerialReviewPending. Exiting to allow other routing types to process.");
                //    return;
                //}

                // Distribution status has to be pending
                if (preDistributionStatus.Value != IsPending)
                {
                    tracer.Trace("Previous Distribution Status was not IsPending. Exiting.");
                    return;
                }

                // Verify completed or rejected
                if (
                    postDistributionStatus.Value != Complete
                    && postDistributionStatus.Value != Rejected
                )
                {
                    tracer.Trace(
                        $"Distribution status changed to {postDistributionStatus.Value}, which is neither Complete nor Rejected. Exiting."
                    );
                    return;
                }

                // Get parent
                var parentReference = postImage.GetAttributeValue<EntityReference>(ParentId);
                if (parentReference == null)
                {
                    throw new Exception(
                        $"Parent routing summary lookup ({ParentId}) missing from distribution."
                    );
                }
                var parentEntity = sysService.Retrieve(
                    ParentEntityName,
                    parentReference.Id,
                    new ColumnSet("ownerid")
                );

                // If rejected
                if (postDistributionStatus.Value == Rejected)
                {
                    // TODO: Should this set the status of remaining reviewers as well?
                    tracer.Trace("Reviewer Rejected. Terminating Workflow.");

                    Entity parentUpdate = new Entity(
                        parentReference.LogicalName,
                        parentReference.Id
                    );
                    parentUpdate[FlowStatus] = new OptionSetValue(WorkflowTerminated);
                    parentUpdate[RoutStatus] = new OptionSetValue(RejectedByReviewer);

                    sysService.Update(parentUpdate);
                    return;
                }

                // If completed
                if (postDistributionStatus.Value == Complete)
                {
                    tracer.Trace("Reviewer Completed. Finding next reviewer.");

                    // Get next reviewer
                    QueryExpression queryNextReviewer = new QueryExpression(ChildEntityName)
                    {
                        ColumnSet = new ColumnSet(DistStatus),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression(
                                    ParentId,
                                    ConditionOperator.Equal,
                                    parentReference.Id
                                ),
                                new ConditionExpression(
                                    DistStatus,
                                    ConditionOperator.Equal,
                                    NotStarted
                                )
                            }
                        }
                    };

                    // Get next reviewer order
                    queryNextReviewer.AddOrder(SetOrder, OrderType.Ascending);
                    queryNextReviewer.TopCount = 1;

                    EntityCollection nextReviewers = sysService.RetrieveMultiple(queryNextReviewer);
                    // Reviewer assigned, reviewer finishes, count starts at 0 each iteration
                    if (nextReviewers.Entities.Count > 0)
                    {
                        // Set next reviewer to IsPending and transfer ownership
                        Entity nextReviewer = nextReviewers.Entities[0];
                        Entity updateReviewer = new Entity(ChildEntityName, nextReviewer.Id);

                        updateReviewer[DistStatus] = new OptionSetValue(IsPending);
                        updateReviewer["ownerid"] = nextReviewer.GetAttributeValue<EntityReference>(
                            ReviewerLookup
                        );

                        sysService.Update(updateReviewer);

                        tracer.Trace("Next reviewer updated to IsPending.");
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

                    // Transfer ownership to parent owner no matter what
                    sysService.Update(
                        new Entity(postImage.LogicalName, postImage.Id)
                        {
                            ["ownerId"] = parentEntity.GetAttributeValue<EntityReference>("ownerid")
                        }
                    );
                    tracer.Trace("Reviewer ownership transferred to parent owner.");
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Error in HandleSerialReviewerProgressPlugin: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}
