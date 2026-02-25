using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DcoumentRouterPlugins.Models
{
    public static class DistributionStatusModel
    {
        public static int NotStarted = 905200000;
        public static int IsPending = 905200001;
        public static int Complete = 905200002;
        public static int Rejected = 905200005;
    }
}
