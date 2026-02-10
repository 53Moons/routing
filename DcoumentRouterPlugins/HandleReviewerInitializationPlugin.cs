
using Microsoft.Xrm.Sdk;

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

            //trigger when status goes from Draft to Pending Review
            // plugin registration is only on status column
            // doesnt mean the status column changed


            // parallel or serial?

            // if parallel, bulk create "Action items"

            // if serial, create first action item (order dependant)
        }
    }
}
