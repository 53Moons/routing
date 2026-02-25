using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Workflow.ComponentModel;

namespace DcoumentRouterPlugins.Models
{
    public static class WorkflowStatusModel
    {
        public static int NotStarted = 905200000;
        public static int InProgress = 905200001;
        public static int Completed = 905200002;
        public static int Terminated = 905200015;
    }
}
