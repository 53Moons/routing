
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


            // parallel or serial?

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
