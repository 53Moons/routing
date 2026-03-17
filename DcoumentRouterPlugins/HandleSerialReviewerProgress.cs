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
        private const int ReviewComplete = 905200002;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Routing Type OptionSet Value
        private const int Serial = 905200000;
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

        // Reviewer Approver lookup fields
        private const string ReviewerLookup = "cr8d2_distributionname";      

        // Owner Email
        private const string OwnerEmail = "cr8d2_owneremail";

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
                    throw new Exception($"Parent routing summary lookup ({ParentId}) missing from reviewer distribution.");
                }

                // Check routing type is serial and get owner email  
                Entity parent = sysService.Retrieve(ParentEntityName, parentReference.Id, new ColumnSet(RoutType, OwnerEmail));
                if (!parent.Contains(RoutType) || parent.GetAttributeValue<OptionSetValue>(RoutType).Value != Serial)

                {
                    tracer.Trace("Routing Type is not Serial. Exiting.");
                    return;
                }

                // If rejected
                if (postDistributionStatus.Value == Rejected)
                {
                    tracer.Trace("Reviewer Rejected. Terminating Workflow.");

                    Entity parentUpdate = new Entity(ParentEntityName, parentReference.Id);
                    parentUpdate[FlowStatus] = new OptionSetValue(WorkflowTerminated);
                    parentUpdate[RoutStatus] = new OptionSetValue(RejectedByReviewer);
                    parentUpdate[ActionWith] = "None";
                    parentUpdate[ActionNext] = "None";

                    sysService.Update(parentUpdate);
                    return;
                }

                // If completed
                if (postDistributionStatus.Value == Complete)
                {
                    tracer.Trace("Reviewer Completed. Finding next reviewer.");

                    // Get next reviewer - updated to pull 2 for action with and action next
                    QueryExpression queryNextReviewer = new QueryExpression(ChildEntityName)
                    {
                        ColumnSet = new ColumnSet(DistStatus, ReviewerLookup),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression(ParentId, ConditionOperator.Equal, parentReference.Id),
                                new ConditionExpression(DistStatus, ConditionOperator.Equal, NotStarted)
                            }
                        }
                    };

                    // Get next reviewer order
                    queryNextReviewer.AddOrder(SetOrder, OrderType.Ascending);
                    queryNextReviewer.TopCount = 2;

                    EntityCollection nextReviewers = sysService.RetrieveMultiple(queryNextReviewer);

                    // Reviewer assigned, reviewer finishes, count starts at 0 each iteration
                    if (nextReviewers.Entities.Count > 0)
                    {
                        // Set next reviewer to IsPending
                        Entity nextReviewer = nextReviewers.Entities[0];
                        Entity updateReviewer = new Entity(ChildEntityName, nextReviewer.Id);

                        updateReviewer[DistStatus] = new OptionSetValue(IsPending);
                        sysService.Update(updateReviewer);

                        tracer.Trace("Next reviewer updated to IsPending.");

                        // Set current reviewer as action with and next reviewer (if exists) as action next
                        EntityReference nextReviewerRef = nextReviewer.GetAttributeValue<EntityReference>(ReviewerLookup);
                        string actionWithName = null;
                        if (nextReviewerRef != null)
                        {
                            actionWithName = nextReviewerRef.Name;
                        }
                        string actionNextName = null;
                        if (nextReviewers.Entities.Count > 1)
                        {
                            EntityReference secondReviewerRef = nextReviewers.Entities[1].GetAttributeValue<EntityReference>(ReviewerLookup);
                            if (secondReviewerRef != null)
                            {
                                actionNextName = secondReviewerRef.Name;
                            }
                            tracer.Trace("Second reviewer found for action next.");
                        }
                        else
                        {
                            tracer.Trace("Second reviewer not found");
                        }

                        Entity parentActionUpdate = new Entity(ParentEntityName, parentReference.Id);
                        parentActionUpdate[ActionWith] = actionWithName;
                        parentActionUpdate[ActionNext] = actionNextName;
                        sysService.Update(parentActionUpdate);

                        tracer.Trace("Parent routing summary updated with action with and action next.");
                    }
                    else
                    {
                        // No additional reviewers found. Review is complete
                        tracer.Trace("No additional reviewers found. Review complete.");

                        string ownerEmail = parent.GetAttributeValue<string>(OwnerEmail);

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
                tracer.Trace($"Error in HandleSerialReviewerProgressPlugin: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
            }



        }
    }
