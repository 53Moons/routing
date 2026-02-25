using DcoumentRouterPlugins.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleReviewerInitializationPlugin : PluginBase
    {
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

                if (
                    !postImage.TryGetAttributeValue(
                        DocumentRouterModel.RoutingStatus,
                        out OptionSetValue postRoutingStatus
                    )
                )
                    throw new Exception("Routing Status not in Post Image");
                if (
                    !preImage.TryGetAttributeValue(
                        DocumentRouterModel.RoutingStatus,
                        out OptionSetValue preRoutingStatus
                    )
                )
                    throw new Exception("Routing Status not in Pre Image");

                if (
                    preRoutingStatus.Value != WorkflowStatusModel.NotStarted
                    || postRoutingStatus.Value != WorkflowStatusModel.InProgress
                )
                {
                    tracer.Trace(
                        $"Routing status changed from {preRoutingStatus.Value} to {postRoutingStatus.Value}. Exiting."
                    );
                    return;
                }
                #endregion

                if (
                    !postImage.TryGetAttributeValue(
                        DocumentRouterModel.RoutingType,
                        out OptionSetValue postRoutingType
                    )
                )
                    throw new Exception("Routing Type not found in Post Image");

                #region Get Reviewers

                var reviewerQuery = new QueryExpression("cr8d2_documentrouterdecision")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression(
                                "cr8d2_routingsummary",
                                ConditionOperator.Equal,
                                postImage.Id
                            ),
                            new ConditionExpression(
                                "statecode",
                                ConditionOperator.Equal,
                                0 // Active
                            )
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
                // if parallel, bulk create "Action items"
                if (postRoutingType.Value == RoutingTypeModel.Parallel)
                {
                    var updates = new EntityCollection();
                    foreach (var reviewer in reviewers.Entities)
                    {
                        reviewer["cr8d2_distributionstatus"] = new OptionSetValue(
                            DistributionStatusModel.IsPending
                        );
                        reviewer["ownerid"] = reviewer.GetAttributeValue<EntityReference>(
                            ReviewerDistributionModel.ReviewerLookup
                        );

                        updates.Entities.Add(reviewer);
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
                }
                #endregion

                #region Handle Serial
                // if serial, create first action item (order dependant)
                else if (postRoutingType.Value == RoutingTypeModel.Serial)
                {
                    var firstReviewer = reviewers.Entities[0];
                    firstReviewer["cr8d2_distributionstatus"] = new OptionSetValue(
                        DistributionStatusModel.IsPending
                    );
                    firstReviewer["ownerid"] = firstReviewer.GetAttributeValue<EntityReference>(
                        ReviewerDistributionModel.ReviewerLookup
                    );

                    try
                    {
                        sysService.Update(firstReviewer);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error creating action item.", ex);
                    }
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
