using DcoumentRouterPlugins.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleSerialReviewerProgressPlugin : PluginBase
    {
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
                        ReviewerDistributionModel.DistStatus,
                        out OptionSetValue postDistributionStatus
                    )
                )
                    throw new Exception("Distribution Status not in Post Image");
                if (
                    !preImage.TryGetAttributeValue(
                        ReviewerDistributionModel.DistStatus,
                        out OptionSetValue preDistributionStatus
                    )
                )
                    throw new Exception("Distribution Status not in Pre Image");

                // Check Workflow Status serial review pending (triggered from initial flow)
                OptionSetValue preWorkflowStatus = preImage.GetAttributeValue<OptionSetValue>(
                    ReviewerDistributionModel.FlowStatus
                );
                OptionSetValue postWorkflowStatus = postImage.GetAttributeValue<OptionSetValue>(
                    ReviewerDistributionModel.FlowStatus
                );

                //bool isSerialReview = (preWorkflowStatus != null && preWorkflowStatus.Value == SerialReviewPending) ||
                //                      (postWorkflowStatus != null && postWorkflowStatus.Value == SerialReviewPending);

                //if (!isSerialReview)
                //{
                //    tracer.Trace("Workflow status is not SerialReviewPending. Exiting to allow other routing types to process.");
                //    return;
                //}

                // Distribution status has to be pending
                if (preDistributionStatus.Value != DistributionStatusModel.IsPending)
                {
                    tracer.Trace("Previous Distribution Status was not IsPending. Exiting.");
                    return;
                }

                // Verify completed or rejected
                if (
                    postDistributionStatus.Value != DistributionStatusModel.Complete
                    && postDistributionStatus.Value != DistributionStatusModel.Rejected
                )
                {
                    tracer.Trace(
                        $"Distribution status changed to {postDistributionStatus.Value}, which is neither Complete nor Rejected. Exiting."
                    );
                    return;
                }

                // Get parent
                var parentReference = postImage.GetAttributeValue<EntityReference>(
                    ReviewerDistributionModel.ParentEntityName
                );
                if (parentReference == null)
                {
                    throw new Exception(
                        $"Parent routing summary lookup ({ReviewerDistributionModel.ParentEntityName}) missing from distribution."
                    );
                }
                var parentEntity = sysService.Retrieve(
                    parentReference.LogicalName,
                    parentReference.Id,
                    new ColumnSet("ownerid")
                );

                // If rejected
                if (postDistributionStatus.Value == DistributionStatusModel.Rejected)
                {
                    // TODO: Should this set the status of remaining reviewers as well?
                    tracer.Trace("Reviewer Rejected. Terminating Workflow.");

                    Entity parentUpdate = new Entity(
                        parentReference.LogicalName,
                        parentReference.Id
                    );
                    parentUpdate[DocumentRouterModel.FlowStatus] = new OptionSetValue(
                        WorkflowStatusModel.Terminated
                    );
                    parentUpdate[DocumentRouterModel.RoutingStatus] = new OptionSetValue(
                        RoutingStatusModel.RejectedByReviewer
                    );

                    sysService.Update(parentUpdate);
                    return;
                }

                // If completed
                if (postDistributionStatus.Value == DistributionStatusModel.Complete)
                {
                    tracer.Trace("Reviewer Completed. Finding next reviewer.");

                    // Get next reviewer
                    QueryExpression queryNextReviewer = new QueryExpression(postImage.LogicalName)
                    {
                        ColumnSet = new ColumnSet(true),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression(
                                    ReviewerDistributionModel.ParentEntityName,
                                    ConditionOperator.Equal,
                                    parentReference.Id
                                ),
                                new ConditionExpression(
                                    ReviewerDistributionModel.DistStatus,
                                    ConditionOperator.Equal,
                                    DistributionStatusModel.NotStarted
                                )
                            }
                        }
                    };

                    // Get next reviewer order
                    queryNextReviewer.AddOrder(
                        ReviewerDistributionModel.SetOrder,
                        OrderType.Ascending
                    );
                    queryNextReviewer.TopCount = 1;

                    EntityCollection nextReviewers = sysService.RetrieveMultiple(queryNextReviewer);
                    // Reviewer assigned, reviewer finishes, count starts at 0 each iteration
                    if (nextReviewers.Entities.Count > 0)
                    {
                        // Set next reviewer to IsPending and transfer ownership
                        Entity nextReviewer = nextReviewers.Entities[0];
                        Entity updateReviewer = new Entity(
                            nextReviewer.LogicalName,
                            nextReviewer.Id
                        );

                        updateReviewer[ReviewerDistributionModel.DistStatus] = new OptionSetValue(
                            DistributionStatusModel.IsPending
                        );
                        updateReviewer["ownerid"] = nextReviewer.GetAttributeValue<EntityReference>(
                            ReviewerDistributionModel.ReviewerLookup
                        );

                        sysService.Update(updateReviewer);

                        tracer.Trace("Next reviewer updated to IsPending.");
                    }
                    else
                    {
                        // No additional reviewers found. Review is complete.
                        tracer.Trace("No additional reviewers. Review complete");

                        Entity parentUpdate = new Entity(
                            parentReference.LogicalName,
                            parentReference.Id
                        );
                        parentUpdate[DocumentRouterModel.RoutingStatus] = new OptionSetValue(
                            RoutingStatusModel.ReviewComplete
                        );
                        parentUpdate[DocumentRouterModel.FlowStatus] = new OptionSetValue(
                            WorkflowStatusModel.PendingInitiatorAction
                        );

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
