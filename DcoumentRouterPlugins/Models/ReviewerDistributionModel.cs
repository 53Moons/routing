using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DcoumentRouterPlugins.Models
{
    public static class ReviewerDistributionModel
    {
        public static string DistStatus = "cr8d2_distributionstatus";
        public static string RoutStatus = "cr8d2_routingstatus";
        public static string FlowStatus = "cr8d2_workflowstatus";

        // Entity References
        public static string ParentEntityName = "cr8d2_routingsummary";
        public static string ChildEntityName = "cr8d2_documentrouterdecision";
        public static string ReviewerLookup = "cr8d2_distributionname";

        // Handle Order
        public static string ParentId = "cr8d2_routingsummaryid";
        public static string SetOrder = "cr8d2_order";
    }
}
