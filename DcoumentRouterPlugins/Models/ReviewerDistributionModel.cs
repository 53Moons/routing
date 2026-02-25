using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DcoumentRouterPlugins.Models
{
    public static class ReviewerDistributionModel
    {
        // Distribution Status OptionSet Values

        public static string DistStatus = "cr8d2_distributionstatus";

        // Routing Status OptionSet Value
        public static int ReviewComplete = 905000002;
        public static string RoutStatus = "cr8d2_routingstatus";

        // Workflow Status OptionSet Value
        public static int PendingInitiatorAction = 905200012;
        public static int SerialReviewPending = 905200003;
        public static string FlowStatus = "cr8d2_workflowstatus";

        // Handle Reject Response
        public static int RejectedByReviewer = 905200006;
        public static int WorkflowTerminated = 905200015;

        // Entity References
        public static string ParentEntityName = "cr8d2_routingsummary";
        public static string ChildEntityName = "cr8d2_documentrouterdecision";
        public static string ReviewerLookup = "cr8d2_distributionname";

        // Handle Order
        public static string ParentId = "cr8d2_routingsummaryid";
        public static string SetOrder = "cr8d2_order";
    }
}
