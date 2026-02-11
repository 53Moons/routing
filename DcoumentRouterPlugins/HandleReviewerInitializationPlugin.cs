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
        private const string ActionTableName = "cr8d2_documentrouteractionitem";

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
                #region Parse RoutingStatus change
                if (!context.PostEntityImages.TryGetValue("Image", out Entity postImage))
                    throw new Exception("Post Image is required.");
                if (!context.PreEntityImages.TryGetValue("Image", out Entity preImage))
                    throw new Exception("Pre Image is required.");

                if (
                    !postImage.TryGetAttributeValue(
                        "cr8d2_routingstatus",
                        out OptionSetValue postRoutingStatus
                    )
                )
                    throw new Exception("Routing Status not in Post Image");
                if (
                    !preImage.TryGetAttributeValue(
                        "cr8d2_routingstatus",
                        out OptionSetValue preRoutingStatus
                    )
                )
                    throw new Exception("Routing Status not in Pre Image");

                if (
                    preRoutingStatus.Value != NotRouted
                    || postRoutingStatus.Value != RoutedForReview
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
                        "cr8d2_routingtype",
                        out OptionSetValue postRoutingType
                    )
                )
                    throw new Exception("Routing Type not found in Post Image");

                #region GetReviewers

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
                if (postRoutingType.Value == Parallel)
                {
                    var actionItems = new EntityCollection();
                    foreach (var reviewer in reviewers.Entities)
                    {
                        var actionItem = new Entity(ActionTableName)
                        {
                            ["cr8d2_reviewer"] = reviewer.ToEntityReference(),
                            ["cr8d2_routingsummary"] = postImage.ToEntityReference()
                        };
                        actionItems.Entities.Add(actionItem);
                    }
                    var createRequest = new CreateMultipleRequest { Targets = actionItems };
                    try
                    {
                        sysService.Execute(createRequest);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error creating action items.", ex);
                    }
                }
                #endregion

                #region Handle Serial
                // if serial, create first action item (order dependant)
                else if (postRoutingType.Value == Serial)
                {
                    var firstReviewer = reviewers.Entities[0];
                    var actionItem = new Entity(ActionTableName)
                    {
                        ["cr8d2_reviewer"] = firstReviewer.ToEntityReference(),
                        ["cr8d2_routingsummary"] = postImage.ToEntityReference()
                    };
                    try
                    {
                        sysService.Create(actionItem);
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
