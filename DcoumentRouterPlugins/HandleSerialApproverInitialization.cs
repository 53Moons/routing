using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class InitializeSerialApproverRoutingPlugin : PluginBase
    {
        // Routing Status OptionSet Values
        private const int ReviewComplete = 905200002;
        private const int RoutedApprover = 905200003;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Routing Type OptionSet Values
        private const int Serial = 905200000;
        private const string RoutType = "cr8d2_routingtype";

        // Approver Reference      
        private const string ParentId = "cr8d2_routingsummary";
        private const string SetOrder = "cr8d2_order";
        private const string ApproverEntityName = "cr8d2_documentroutermanagerdistribution";

        // Distribution Status OptionSet Values
        private const int IsPending = 905200001;
        private const string DistStatus = "cr8d2_distributionstatus";

        public InitializeSerialApproverRoutingPlugin() 
            : base(typeof(InitializeSerialApproverRoutingPlugin)) {
           
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            tracer.Trace("Start InitializeSerialApproverRoutingPlugin");

            if (context.MessageName != "Update" || context.Stage != 40)
                return;
            try
            {

                if (!context.PostEntityImages.TryGetValue("PostImage", out Entity postImage))
                    throw new InvalidPluginExecutionException("PostImage is required.");
                if (!context.PreEntityImages.TryGetValue("PreImage", out Entity preImage))
                    throw new InvalidPluginExecutionException("PreImage is required.");

                // Routing Status should change from ReviewComplete to RoutedApprover
                var preRoutStatus = preImage.GetAttributeValue<OptionSetValue>(RoutStatus);
                var postRoutStatus = postImage.GetAttributeValue<OptionSetValue>(RoutStatus);

                if (preRoutStatus == null || postRoutStatus == null ||
                    preRoutStatus.Value != ReviewComplete ||
                    postRoutStatus.Value != RoutedApprover)
                {
                    tracer.Trace("Routing status did not change to RoutedApprover. Exiting.");
                    return;
                }

                // Routing Type should be Serial
                var postRoutType = postImage.GetAttributeValue<OptionSetValue>(RoutType);
                if (postRoutType == null || postRoutType.Value != Serial)
                {
                    tracer.Trace("Routing type is not Serial. Exiting.");
                    return;
                }

                // Find first approver
                var approverQuery = new QueryExpression(ApproverEntityName)
                {
                    ColumnSet = new ColumnSet(DistStatus, SetOrder),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                    {
                        new ConditionExpression(ParentId, ConditionOperator.Equal, postImage.Id),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active
                    }
                    },
                    TopCount = 1
                };
                approverQuery.AddOrder(SetOrder, OrderType.Ascending);

                var approvers = sysService.RetrieveMultiple(approverQuery);
                if (approvers.Entities.Count == 0)
                {
                    tracer.Trace("No approvers found for this document.");
                    return;
                }

                // Update first approver to IsPending
                var firstApprover = approvers.Entities[0];

                Entity updateApprover = new Entity(ApproverEntityName, firstApprover.Id);
                updateApprover[DistStatus] = new OptionSetValue(IsPending);

                sysService.Update(updateApprover);
                tracer.Trace("First approver set to IsPending successfully.");
            }
            catch (Exception ex)
            {
                tracer.Trace($"Error in InitializeSerialApproverRoutingPlugin: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}