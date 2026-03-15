using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleReviewerInitializationPlugin : PluginBase
    {
        private const int RoutedForReview = 905200001;
        private const int NotRouted = 905200000;
        private const int Serial = 905200000;
        private const int Parallel = 905200001;
        private const int IsPending = 905200001;

        // Routing summary fields to set actionwith and actionnext
        private const string ActionWith = "cr8d2_actionwith";
        private const string ActionNext = "cr8d2_actionnext";

        // Reviewer name lookup field on routing decision entity
        private const string ReviewerLookup = "cr8d2_distributionname";

        public HandleReviewerInitializationPlugin()
            : base(typeof(HandleReviewerInitializationPlugin))
        {
            // Not Implemented
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            if (context.MessageName != "Update" || context.Stage != 40)
                return;
            try
            {
                #region Parse Routing Status change
                if (!context.PostEntityImages.TryGetValue("Image", out Entity postImage))
                    throw new Exception("Post Image is required.");
                if (!context.PreEntityImages.TryGetValue("Image", out Entity preImage))
                    throw new Exception("Pre Image is required.");

                if (!postImage.TryGetAttributeValue("cr8d2_routingstatus", out OptionSetValue postRoutingStatus))
                    throw new Exception("Routing Status not in Post Image");
                if (!preImage.TryGetAttributeValue("cr8d2_routingstatus", out OptionSetValue preRoutingStatus))
                    throw new Exception("Routing Status not in Pre Image");

                if (preRoutingStatus.Value != NotRouted || postRoutingStatus.Value != RoutedForReview)
                {
                    tracer.Trace($"Routing status changed from {preRoutingStatus.Value} to {postRoutingStatus.Value}. Exiting.");
                    return;
                }
                #endregion

                if (!postImage.TryGetAttributeValue("cr8d2_routingtype", out OptionSetValue postRoutingType))
                    throw new Exception("Routing Type not found in Post Image");

                #region Get Reviewers
                var reviewerQuery = new QueryExpression("cr8d2_documentrouterdecision")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("cr8d2_routingsummary", ConditionOperator.Equal, postImage.Id),
                            new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                        }
                    },
                    Orders = { new OrderExpression("cr8d2_order", OrderType.Ascending) }
                };

                EntityCollection reviewers;
                try
                {
                    reviewers = sysService.RetrieveMultiple(reviewerQuery);
                    if (reviewers.Entities.Count == 0)
                    {
                        tracer.Trace("No reviewers found for this document.");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    tracer.Trace($"Error retrieving reviewers: {ex.Message}");
                    throw new Exception("Error retrieving reviewers.", ex);
                }
                #endregion

                #region Handle Parallel
                if (postRoutingType.Value == Parallel)
                {
                    var updates = new EntityCollection { EntityName = "cr8d2_documentrouterdecision" };
                    var reviewerNames = new System.Collections.Generic.List<string>();

                    foreach (var reviewer in reviewers.Entities)
                    {
                        reviewer["cr8d2_distributionstatus"] = new OptionSetValue(IsPending);
                        updates.Entities.Add(reviewer);

                        // Get reviewer for action with
                        EntityReference reviewerRef = reviewer.GetAttributeValue<EntityReference>(ReviewerLookup);
                        if (reviewerRef != null)
                        {
                            reviewerNames.Add(reviewerRef.Name);
                        }
                    }

                    var updateRequest = new UpdateMultipleRequest { Targets = updates };
                    try
                    {
                        sysService.Execute(updateRequest);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error creating action items.", ex);
                    }

                    // parallel action with all names next action owner by name not email so we don't have to query user table
                    EntityReference ownerRef = postImage.GetAttributeValue<EntityReference>("ownerid");

                    string ownerName = null;
                    if (ownerRef != null)
                    {
                        ownerName = ownerRef.Name;
                    }

                    Entity parallelParentUpdate = new Entity("cr8d2_routingsummary", postImage.Id);
                    parallelParentUpdate[ActionWith] = string.Join(", ", reviewerNames);
                    parallelParentUpdate[ActionNext] = ownerName;
                    sysService.Update(parallelParentUpdate);

                    tracer.Trace($"Parallel ActionWith set to: {string.Join(", ", reviewerNames)}");
                }
                #endregion

                #region Handle Serial
                else if (postRoutingType.Value == Serial)
                {
                    var firstReviewer = reviewers.Entities[0];
                    firstReviewer["cr8d2_distributionstatus"] = new OptionSetValue(IsPending);

                    try
                    {
                        sysService.Update(firstReviewer);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error creating action item.", ex);
                    }

                    // Set action with to first reviewer and action next to second if exists
                    EntityReference firstReviewerRef = firstReviewer.GetAttributeValue<EntityReference>(ReviewerLookup);

                    string actionWithName = null;
                    if (firstReviewerRef != null)
                    {
                        actionWithName = firstReviewerRef.Name;
                    }

                    string actionNextName = null;
                    if (reviewers.Entities.Count > 1)
                    {
                        EntityReference secondReviewerRef = reviewers.Entities[1].GetAttributeValue<EntityReference>(ReviewerLookup);
                        if (secondReviewerRef != null)
                        {
                            actionNextName = secondReviewerRef.Name;
                        }
                        tracer.Trace("ActionNext set to second reviewer.");
                    }
                    else
                    {
                        tracer.Trace("Only one reviewer. ActionNext will be null.");
                    }

                    Entity serialParentUpdate = new Entity("cr8d2_routingsummary", postImage.Id);
                    serialParentUpdate[ActionWith] = actionWithName;
                    serialParentUpdate[ActionNext] = actionNextName;
                    sysService.Update(serialParentUpdate);

                    tracer.Trace("ActionWith and ActionNext set on routing summary.");
                }
                #endregion
            }
            catch (Exception ex)
            {
                tracer.Trace($"Unhandled exception: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}