using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace DcoumentRouterPlugins
{
    public class HandleReviewOrderPlugin : PluginBase
    {
        private readonly string _orderColumnName;
        private readonly string _lookupColumnName;

        public HandleReviewOrderPlugin(string unsecureConfig, string _)
            : base(typeof(HandleReviewOrderPlugin))
        {
            string[] unsecureConfigs = unsecureConfig.Split(',');
            _orderColumnName = unsecureConfigs[0];
            _lookupColumnName = unsecureConfigs[1];
        }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            if (string.IsNullOrEmpty(_orderColumnName))
                throw new InvalidPluginExecutionException(
                    "Invalid Registration. Missing order column name."
                );

            if (string.IsNullOrEmpty(_lookupColumnName))
                throw new InvalidPluginExecutionException(
                    "Invalid Registration. Missing lookup column name."
                );

            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.CurrentUserService;
            var tracer = localPluginContext.TracingService;

            // Post-operation only
            if (context.Stage != 40)
                return;

            if (context.MessageName != "Create" && context.MessageName != "Update")
                throw new InvalidPluginExecutionException(
                    "Invalid Registration. HandleReviewOrderPlugin on handles Create and Update"
                );

            if (!context.InputParameters.TryGetValue("Target", out Entity target))
                throw new InvalidPluginExecutionException("Target missing");

            tracer.Trace($"Target {target.LogicalName} ({target.Id})");

            if (!target.TryGetAttributeValue(_orderColumnName, out int order))
                throw new InvalidPluginExecutionException("Order property not found on value");

            tracer.Trace($"Order {order}");
            int? oldOrder = null;

            if (context.MessageName == "Update")
            {
                if (!context.PreEntityImages.TryGetValue("PreImage", out Entity preImage))
                    throw new InvalidPluginExecutionException("PreImage required on Update");

                if (!preImage.TryGetAttributeValue(_orderColumnName, out int preOrder))
                    throw new InvalidPluginExecutionException("PreImage missing order column.");

                oldOrder = preOrder;
            }

            if (target.TryGetAttributeValue(_lookupColumnName, out EntityReference lookup))
            {
                // ok
            }
            else if (context.MessageName == "Create")
                // Create requires lookup
                throw new InvalidPluginExecutionException(
                    $"{_lookupColumnName} required on Create"
                );
            else
            {
                // No Lookup in target on update
                // Fall back to PreImage
                if (!context.PreEntityImages.TryGetValue("PreImage", out Entity preImage))
                    throw new InvalidPluginExecutionException("PreImage required on Update");

                if (!preImage.TryGetAttributeValue(_lookupColumnName, out lookup))
                    throw new InvalidPluginExecutionException(
                        $"{_lookupColumnName} value not found on target or PreImage."
                    );
            }

            tracer.Trace($"Lookup {lookup.LogicalName} ({lookup.Id})");

            // If Create, treat as "insert at order", shift everything >= order up by 1
            // If Update, shift only the affected range depending on direction
            FilterExpression criteria = new FilterExpression(LogicalOperator.And);

            criteria.Conditions.Add(
                new ConditionExpression(
                    target.LogicalName + "id",
                    ConditionOperator.NotEqual,
                    target.Id
                )
            );
            criteria.Conditions.Add(
                new ConditionExpression(_lookupColumnName, ConditionOperator.Equal, lookup.Id)
            );

            if (context.MessageName == "Create")
            {
                // shift all >= new order up
                criteria.Conditions.Add(
                    new ConditionExpression(_orderColumnName, ConditionOperator.GreaterEqual, order)
                );
            }
            else
            {
                // Update: only shift within the impacted band
                if (!oldOrder.HasValue)
                    return;

                if (order < oldOrder.Value)
                {
                    // moved UP: shift [newOrder .. oldOrder-1] up (+1)
                    criteria.Conditions.Add(
                        new ConditionExpression(
                            _orderColumnName,
                            ConditionOperator.Between,
                            new object[] { order, oldOrder.Value - 1 }
                        )
                    );
                }
                else if (order > oldOrder.Value)
                {
                    // moved DOWN: shift [oldOrder+1 .. newOrder] down (-1)
                    criteria.Conditions.Add(
                        new ConditionExpression(
                            _orderColumnName,
                            ConditionOperator.Between,
                            new object[] { oldOrder.Value + 1, order }
                        )
                    );
                }
                else
                {
                    // no move
                    return;
                }
            }

            var query = new QueryExpression(target.LogicalName)
            {
                ColumnSet = new ColumnSet(_orderColumnName, "createdon"), // keep it light
                Criteria = criteria,
                Orders =
                {
                    new OrderExpression(_orderColumnName, OrderType.Ascending),
                    new OrderExpression("createdon", OrderType.Ascending)
                }
            };

            EntityCollection results;
            try
            {
                results = service.RetrieveMultiple(query);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

            tracer.Trace($"Related record count {results.Entities.Count}");

            if (results.Entities.Count == 0)
                return;

            var updatedCollection = ReorderAffectedRange(
                results,
                _orderColumnName,
                oldOrder,
                order
            );

            var request = new UpdateMultipleRequest() { Targets = updatedCollection };
            request.Parameters.Add("BypassBusinessLogicExecution", "CustomSync,CustomAsync");

            try
            {
                service.Execute(request);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Failed to execute {ex.Message}");
            }
        }

        private EntityCollection ReorderAffectedRange(
            EntityCollection collection,
            string columnName,
            int? oldOrder,
            int newOrder
        )
        {
            var updatedCollection = new EntityCollection() { EntityName = collection.EntityName };

            if (collection == null || collection.Entities == null || collection.Entities.Count == 0)
                return updatedCollection;

            // CREATE: shift all >= newOrder up by +1
            if (!oldOrder.HasValue)
            {
                foreach (var entity in collection.Entities)
                {
                    var current = entity.GetAttributeValue<int?>(columnName) ?? 0;
                    updatedCollection.Entities.Add(
                        new Entity(entity.LogicalName)
                        {
                            Id = entity.Id,
                            [columnName] = current + 1
                        }
                    );
                }

                return updatedCollection;
            }

            // UPDATE: direction depends on where the record moved
            int direction;
            if (newOrder < oldOrder.Value)
                direction = +1; // moved up
            else if (newOrder > oldOrder.Value)
                direction = -1; // moved down
            else
                return updatedCollection;

            foreach (var entity in collection.Entities)
            {
                var current = entity.GetAttributeValue<int?>(columnName) ?? 0;
                updatedCollection.Entities.Add(
                    new Entity(entity.LogicalName)
                    {
                        Id = entity.Id,
                        [columnName] = current + direction
                    }
                );
            }

            return updatedCollection;
        }
    }
}
