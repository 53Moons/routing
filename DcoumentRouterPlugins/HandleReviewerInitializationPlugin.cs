using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
﻿
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;

namespace DcoumentRouterPlugins
{
    public class HandleReviewerInitializationPlugin : PluginBase
    {
        private const int RoutedForReview = 905200001;
        private const int NotRouted = 905200000;
        private const int Serial = 905200000;
        private const int Parallel = 905200001;
        private const int IsPending = 905200001;

        public HandleReviewerInitializationPlugin()
            : base(typeof(HandleReviewerInitializationPlugin))
        {
            // Not Implemented
        }
        public new void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context =
                (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory serviceFactory =
            (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            
            serviceFactory.CreateOrganizationService(context.UserId);
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
            //trigger when status goes from Draft to Pending Review
            if (context.MessageName != "Update" || context.PrimaryEntityName != "cr8d2_routingsummary")
            {
                return;
            }
            if (!context.InputParameters.Contains("Target"))
            {
                return;
            }
            Entity target = (Entity)context.InputParameters["Target"];
            
            if (!target.Contains("routingstatus"))
            {
                return;
            }

            Entity preImage = context.PreEntityImages["PreImage"];
            Entity postImage = context.PostEntityImages["PostImage"];
            

            OptionSetValue preStatus = preImage.GetAttributeValue<OptionSetValue>("routingstatus");
            OptionSetValue postStatus = postImage.GetAttributeValue<OptionSetValue>("routingstatus");

            const int PendingReview = 905200001;

            if (postStatus.Value != PendingReview)
            {
                return;
            }
            Guid parentId = postImage.Id;

            // plugin registration is only on status column
            // doesnt mean the status column changed

                #region Handle Parallel
                // if parallel, bulk create "Action items"
                if (postRoutingType.Value == Parallel)
                {
                    var updates = new EntityCollection();
                    foreach (var reviewer in reviewers.Entities)
                    {
                        reviewer["cr8d2_distributionstatus"] = new OptionSetValue(IsPending);

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
                }
                #endregion
            }
            catch (Exception ex)
            {
                tracer.Trace($"Unhandled exception: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            OptionSetValue preStatusType = preImage.GetAttributeValue<OptionSetValue>("routingtype");
            OptionSetValue postStatusType = postImage.GetAttributeValue<OptionSetValue>("routingtype");

            const int RoutingType_Serial = 905200000;
            const int RoutingType_Parallel = 905200001;

            if (postStatusType.Value == RoutingType_Serial)
            {
                return;
            }
        }
            public static EntityCollection SimpleExample(IOrganizationService service) {
                QueryExpression query = new QueryExpression("cr8d2_documentrouterdecision");
                query.ColumnSet = new ColumnSet();
                query.ColumnSet.AddColumns("cr8d2_name", "cr8d2_order", "cr8d2_distributionstatus");
                query.AddOrder("cr8d2_order", OrderType.Ascending);
                query.Criteria.AddCondition("cr8d2_routingsummaryid", ConditionOperator.Equal, parentId);
                return service.RetrieveMultiple(query);
            }

        // if parallel, bulk create "Action items"

        // if serial, create first action item (order dependant)
    }
}
