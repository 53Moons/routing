
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
        }
    }
}
