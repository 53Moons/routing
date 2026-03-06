using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleNewParallelReviewerCreatePlugin : PluginBase
    {
        // Routing Status OptionSet Values
        private const int RoutedForReview = 905200001;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Routing Type OptionSet Values
        private const int Parallel = 905200001;
        private const string RoutType = "cr8d2_routingtype";

        // Distribution Status OptionSet Values
        private const int IsPending = 905200001;
        private const string DistStatus = "cr8d2_distributionstatus";

        // Entity References
        private const string ParentEntityName = "cr8d2_routingsummary";
        private const string ChildEntityName = "cr8d2_documentrouterdecision";


        public HandleNewParallelReviewerCreatePlugin()
            : base(typeof(HandleNewParallelReviewerCreatePlugin))
        {
            // Not Implemented
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            // Check stage is pre operation 20 
            if (context.MessageName != "Create" || context.Stage != 20)
                return;

            try
            {
                if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity targetEntity))
                    throw new Exception("Target entity is missing.");

                // Get parent reference
                var parentReference = targetEntity.GetAttributeValue<EntityReference>(ParentEntityName);
                if (parentReference == null)
                {
                    tracer.Trace("No parent found. Exiting.");
                    return;
                }

                // Verify routing status and routing type
                Entity parentDocument = sysService.Retrieve(parentReference.LogicalName,parentReference.Id,
                    new ColumnSet(RoutStatus, RoutType)
                );

                var parentRoutingStatus = parentDocument.GetAttributeValue<OptionSetValue>(RoutStatus)?.Value;
                var parentRoutingType = parentDocument.GetAttributeValue<OptionSetValue>(RoutType)?.Value;

                // Confirm routed for review and parallel routing type  
                if (parentRoutingStatus == RoutedForReview && parentRoutingType == Parallel)
                {
                    tracer.Trace("Parent is currently in a Parallel review. Setting new reviewer to IsPending.");

                    // Update distribution status  
                    targetEntity[DistStatus] = new OptionSetValue(IsPending);
                }
                else
                {
                    tracer.Trace("Parent is either not routing or not parallel. Reviewer defaults will apply.");
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Unhandled exception in HandleNewParallelReviewerCreatePlugin: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}