using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DcoumentRouterPlugins.Models
{
    public static class RoutingStatusModel
    {
        public static int NotRouted = 905200000;
        public static int RoutedForReview = 905200001;
        public static int ReviewComplete = 905200002;
        public static int RoutedToApprover = 905200003;
        public static int RoutingComplete = 905200004;
        public static int RejectedByApprover = 905200005;
        public static int RejectedByReviewer = 905200006;
    }
}
