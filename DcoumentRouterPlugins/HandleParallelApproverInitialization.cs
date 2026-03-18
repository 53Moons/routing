using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    public class HandleParallelApproverInitialization : PluginBase
    {
        // Distribution Status OptionSet Values
        private const int NotStarted = 905200000;
        private const int IsPending = 905200001;
        private const int Complete = 905200002;
        private const int Rejected = 905200005;
        private const string DistStatus = "cr8d2_distributionstatus";

        // Routing Status OptionSet Value
        private const int ReviewComplete = 905200002;
        private const int RoutedApprover = 905200003;
        private const string RoutStatus = "cr8d2_routingstatus";

        // Routing Type OptionSet Value
        private const int Parallel = 905200001;
        private const string RoutType = "cr8d2_routingtype";

        // Workflow Status OptionSet Value
        private const int PendingInitiatorAction = 905200012;
        private const int FinalApprovalPending = 905200013;
        private const string FlowStatus = "cr8d2_workflowstatus";

        // Entity References
        private const string ParentEntityName = "cr8d2_routingsummary";
        private const string ChildEntityName = "cr8d2_documentroutermanagerdistribution";
        private const string ParentId = "cr8d2_routingsummary";

        // Routing summary fields to set actionwith and actionnext
        private const string ActionWith = "cr8d2_actionwith";
        private const string ActionNext = "cr8d2_actionnext";

        // Reviewer lookup field
        private const string ApproverLookup = "cr8d2_managername";

        // Owner Email
        private const string OwnerEmail = "cr8d2_owneremail";

        public HandleParallelApproverInitialization()
            : base(typeof(HandleParallelApproverInitialization))
        {
            // Not Implemented
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            tracer.Trace("StartParallelApproverInitialization");

            // Check stage is post operation 40
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

                // Routing Type should be Parallel get owner email
                Entity parent = sysService.Retrieve(ParentEntityName, postImage.Id, new ColumnSet(RoutType, OwnerEmail));
                if (!parent.Contains(RoutType) || parent.GetAttributeValue<OptionSetValue>(RoutType).Value != Parallel)
                {
                    tracer.Trace("Routing Type is not Parallel. Exiting.");
                    return;
                }

                // Get child dist records
                var approverQuery = new QueryExpression(ChildEntityName)
                {
                    ColumnSet = new ColumnSet(DistStatus, ApproverLookup),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression(ParentId,ConditionOperator.Equal,postImage.Id),
                            new ConditionExpression("statecode",ConditionOperator.Equal,0) // Active
                        }
                    }
                };

                // Get approvers
                var approvers= sysService.RetrieveMultiple(approverQuery);

                if (approvers.Entities.Count > 0)
                {
                    var updates = new EntityCollection { EntityName = ChildEntityName };
                    System.Collections.Generic.List<string> approverNames = new System.Collections.Generic.List<string>();

                    foreach (var approver in approvers.Entities)
                    {
                        approver[DistStatus] = new OptionSetValue(IsPending);
                        updates.Entities.Add(approver);

                        EntityReference approverRef = approver.GetAttributeValue<EntityReference>(ApproverLookup);
                        if (approverRef != null)
                        {
                            approverNames.Add(approverRef.Name);
                        }
                    }

                    var updateRequest = new UpdateMultipleRequest { Targets = updates };
                    try
                    {
                        sysService.Execute(updateRequest);
                        tracer.Trace($"Successfully updated {updates.Entities.Count} approver distribution records to IsPending.");

                        // Set action with to approvers and action next to owner email
                        string ownerEmail =parent.GetAttributeValue<string>(OwnerEmail);
                        Entity parentUpdate = new Entity(ParentEntityName, postImage.Id);
                        parentUpdate[ActionWith] = string.Join(",", approverNames);
                        parentUpdate[ActionNext] = ownerEmail;

                        tracer.Trace($"ActionWith set to: {string.Join(",", approverNames)}");
                        sysService.Update(parentUpdate);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error updating distribution records.", ex);
                    }
                }
                else
                {
                    tracer.Trace("No active approvers found to initialize.");
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Error in HandleParallelApproverInitialization: {ex.Message}");
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}