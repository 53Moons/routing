using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System;

namespace DcoumentRouterPlugins
{
    public class ManageRoutingUsersTeamMembership : PluginBase
    {
        // Entity reference and lookups
        private const string RoutingUsers = "cr8d2_routinguser";
        private const string RoutingUserLookup = "cr8d2_routingusername";
        private const string TeamLookup = "cr8d2_team";

        public ManageRoutingUsersTeamMembership() : base(typeof(ManageRoutingUsersTeamMembership)) { }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            // Verify table
            if (context.PrimaryEntityName != RoutingUsers) return;

            try
            {
                if (context.MessageName == "Create" && context.Stage == 20)
                {
                    // Stage 20 pre-operation set the team
                    AutoSetTeam(context, sysService, tracer);
                }
                else if (context.MessageName == "Create" && context.Stage == 40)
                {
                    // Stage 40 post-operation add user to team
                    AddUserToTeam(context, sysService, tracer);
                }
                else if (context.MessageName == "Delete" && context.Stage == 40)
                {
                    // Stage 40 post-operation remove user from team
                    RemoveUserFromTeam(context, sysService, tracer);
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Error in ManageRoutingUsers: {ex.Message}");
            }
        }

        private void AutoSetTeam(IPluginExecutionContext context, IOrganizationService sysService, ITracingService tracer)
        {
            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity target)) return;

            QueryExpression query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid")
            };
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "Routing Users");

            var results = sysService.RetrieveMultiple(query);

            if (results.Entities.Count > 0)
            {
                Guid teamId = results.Entities[0].Id;
                target[TeamLookup] = new EntityReference("team", teamId);
                tracer.Trace($"Successfully set the team lookup to Routing Users ({teamId}).");
            }
            else
            {
                throw new InvalidPluginExecutionException("The 'Routing Users' team could not be found.");
            }
        }

        private void AddUserToTeam(IPluginExecutionContext context, IOrganizationService sysService, ITracingService tracer)
        {
            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity target)) return;

            if (!target.Contains(TeamLookup) || !target.Contains(RoutingUserLookup)) return;

            Guid targetTeamId = target.GetAttributeValue<EntityReference>(TeamLookup).Id;
            Guid targetUserId = target.GetAttributeValue<EntityReference>(RoutingUserLookup).Id;

            // Execute message to ADD to the team
            AddMembersTeamRequest addRequest = new AddMembersTeamRequest
            {
                TeamId = targetTeamId,
                MemberIds = new[] { targetUserId }
            };

            sysService.Execute(addRequest);
            tracer.Trace($"Successfully added User {targetUserId} to Team {targetTeamId}.");
        }

        private void RemoveUserFromTeam(IPluginExecutionContext context, IOrganizationService sysService, ITracingService tracer)
        {
            // Deleted user must check PreEntityImage to see the lookups BEFORE the record was deleted
            if (!context.PreEntityImages.Contains("PreImage") || !(context.PreEntityImages["PreImage"] is Entity preImage))
            {
                throw new InvalidPluginExecutionException("Plugin configuration error: A PreEntityImage named 'PreImage' is required to remove the user from the team.");
            }

            if (!preImage.Contains(TeamLookup) || !preImage.Contains(RoutingUserLookup)) return;

            Guid targetTeamId = preImage.GetAttributeValue<EntityReference>(TeamLookup).Id;
            Guid targetUserId = preImage.GetAttributeValue<EntityReference>(RoutingUserLookup).Id;

            // Execute message to REMOVE from the team
            RemoveMembersTeamRequest removeRequest = new RemoveMembersTeamRequest
            {
                TeamId = targetTeamId,
                MemberIds = new[] { targetUserId }
            };

            sysService.Execute(removeRequest);
            tracer.Trace($"Successfully removed User {targetUserId} from Team {targetTeamId}.");
        }
    }
}