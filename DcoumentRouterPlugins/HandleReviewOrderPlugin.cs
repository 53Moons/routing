using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace DcoumentRouterPlugins
{
    /// <summary>
    /// Plugin that automatically manages the ordering of related records.
    /// When a record is created or its order is updated, this plugin shifts other 
    /// records to maintain a proper sequence without gaps or duplicates.
    /// 
    /// Example: If you have records ordered 1, 2, 3, 4 and you insert a new record 
    /// at position 2, the existing records 2, 3, 4 will be shifted to 3, 4, 5.
    /// </summary>
    public class HandleReviewOrderPlugin : PluginBase
    {
        // The name of the column that stores the order/sequence number (e.g., "revieworder")
        private readonly string _orderColumnName;

        // The name of the lookup column that groups related records (e.g., "documentid")
        private readonly string _lookupColumnName;

        /// <summary>
        /// Initializes the plugin with configuration values.
        /// </summary>
        /// <param name="unsecureConfig">
        /// Configuration string in format: "orderColumnName,lookupColumnName"
        /// Example: "revieworder,cr2b3_documentid"
        /// </param>
        /// <param name="_">Secure config (not used in this plugin)</param>
        public HandleReviewOrderPlugin(string unsecureConfig, string _)
            : base(typeof(HandleReviewOrderPlugin))
        {
            // Parse the configuration string to get the column names
            string[] unsecureConfigs = unsecureConfig.Split(',');
            _orderColumnName = unsecureConfigs[0];
            _lookupColumnName = unsecureConfigs[1];
        }

        /// <summary>
        /// Main plugin execution logic. Handles the reordering of records when a record 
        /// is created or updated.
        /// </summary>
        /// <param name="localPluginContext">The plugin context containing execution details</param>
        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            // Validate that required configuration was provided
            if (string.IsNullOrEmpty(_orderColumnName))
                throw new InvalidPluginExecutionException(
                    "Invalid Registration. Missing order column name."
                );

            if (string.IsNullOrEmpty(_lookupColumnName))
                throw new InvalidPluginExecutionException(
                    "Invalid Registration. Missing lookup column name."
                );

            // Get the execution context, service, and tracing service
            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.CurrentUserService;
            var tracer = localPluginContext.TracingService;

            // Only run in post-operation stage (after the record is saved to the database)
            // Stage 40 = Post-operation
            if (context.Stage != 40)
                return;

            // Verify this plugin is registered on Create or Update messages only
            if (context.MessageName != "Create" && context.MessageName != "Update")
                throw new InvalidPluginExecutionException(
                    "Invalid Registration. HandleReviewOrderPlugin on handles Create and Update"
                );

            // Get the record being created or updated (the "Target")
            if (!context.InputParameters.TryGetValue("Target", out Entity target))
                throw new InvalidPluginExecutionException("Target missing");

            tracer.Trace($"Target {target.LogicalName} ({target.Id})");

            // Get the new order value from the target record
            if (!target.TryGetAttributeValue(_orderColumnName, out int order))
                throw new InvalidPluginExecutionException("Order property not found on value");

            tracer.Trace($"Order {order}");
            int? oldOrder = null;

            // For Update operations, we need to know what the old order was
            // This helps us determine which records need to be shifted
            if (context.MessageName == "Update")
            {
                // Get the record's state before the update (PreImage)
                if (!context.PreEntityImages.TryGetValue("PreImage", out Entity preImage))
                    throw new InvalidPluginExecutionException("PreImage required on Update");

                if (!preImage.TryGetAttributeValue(_orderColumnName, out int preOrder))
                    throw new InvalidPluginExecutionException("PreImage missing order column.");

                oldOrder = preOrder;
            }

            // Get the lookup reference that groups related records together
            // For example, all reviews for the same document would share the same lookup
            if (target.TryGetAttributeValue(_lookupColumnName, out EntityReference lookup))
            {
                // Lookup found in the target - we're good
            }
            else if (context.MessageName == "Create")
                // On Create, the lookup is required (we need to know which group this belongs to)
                throw new InvalidPluginExecutionException(
                    $"{_lookupColumnName} required on Create"
                );
            else
            {
                // On Update, if the lookup isn't being changed, get it from the PreImage
                // This handles cases where only the order is being updated
                if (!context.PreEntityImages.TryGetValue("PreImage", out Entity preImage))
                    throw new InvalidPluginExecutionException("PreImage required on Update");

                if (!preImage.TryGetAttributeValue(_lookupColumnName, out lookup))
                    throw new InvalidPluginExecutionException(
                        $"{_lookupColumnName} value not found on target or PreImage."
                    );
            }

            tracer.Trace($"Lookup {lookup.LogicalName} ({lookup.Id})");

            // Build a query to find all other records in the same group that need reordering
            // We'll only retrieve records that are affected by this change
            FilterExpression criteria = new FilterExpression(LogicalOperator.And);

            // Exclude the current record (we don't want to update it again)
            criteria.Conditions.Add(
                new ConditionExpression(
                    target.LogicalName + "id",
                    ConditionOperator.NotEqual,
                    target.Id
                )
            );

            // Only get records in the same group (same lookup value)
            criteria.Conditions.Add(
                new ConditionExpression(_lookupColumnName, ConditionOperator.Equal, lookup.Id)
            );

            if (context.MessageName == "Create")
            {
                // CREATE scenario: A new record is being inserted at position 'order'
                // Shift all records with order >= new position up by 1
                // Example: Inserting at 2 means records 2,3,4 become 3,4,5
                criteria.Conditions.Add(
                    new ConditionExpression(_orderColumnName, ConditionOperator.GreaterEqual, order)
                );
            }
            else
            {
                // UPDATE scenario: A record moved from oldOrder to newOrder
                // Only shift records in the affected range
                if (!oldOrder.HasValue)
                    return;

                if (order < oldOrder.Value)
                {
                    // Record moved UP (to a lower number)
                    // Example: Moving from 5 to 2 means records 2,3,4 shift down (become 3,4,5)
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
                    // Record moved DOWN (to a higher number)
                    // Example: Moving from 2 to 5 means records 3,4,5 shift up (become 2,3,4)
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
                    // Order didn't change - nothing to do
                    return;
                }
            }

            // Build the query to retrieve affected records
            var query = new QueryExpression(target.LogicalName)
            {
                ColumnSet = new ColumnSet(_orderColumnName, "createdon"), // Only get what we need
                Criteria = criteria,
                Orders =
                {
                    // Order results by sequence number, then creation date for consistency
                    new OrderExpression(_orderColumnName, OrderType.Ascending),
                    new OrderExpression("createdon", OrderType.Ascending)
                }
            };

            // Execute the query to get all affected records
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

            // If no records need to be updated, we're done
            if (results.Entities.Count == 0)
                return;

            // Calculate the new order values for all affected records
            var updatedCollection = ReorderAffectedRange(
                results,
                _orderColumnName,
                oldOrder,
                order
            );

            // Update all affected records in a single batch operation
            var request = new UpdateMultipleRequest() { Targets = updatedCollection };

            // Bypass business logic to prevent infinite loops (this plugin won't trigger itself)
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

        /// <summary>
        /// Calculates new order values for all affected records based on the operation type.
        /// </summary>
        /// <param name="collection">The collection of records that need reordering</param>
        /// <param name="columnName">The name of the order column</param>
        /// <param name="oldOrder">The old order value (null for Create operations)</param>
        /// <param name="newOrder">The new order value</param>
        /// <returns>A collection of entities with updated order values ready to be saved</returns>
        private EntityCollection ReorderAffectedRange(
            EntityCollection collection,
            string columnName,
            int? oldOrder,
            int newOrder
        )
        {
            var updatedCollection = new EntityCollection() { EntityName = collection.EntityName };

            // Safety check - if no records to update, return empty collection
            if (collection == null || collection.Entities == null || collection.Entities.Count == 0)
                return updatedCollection;

            // CREATE scenario: shift all records up by 1
            if (!oldOrder.HasValue)
            {
                foreach (var entity in collection.Entities)
                {
                    var current = entity.GetAttributeValue<int?>(columnName) ?? 0;

                    // Create a new entity with just the ID and updated order value
                    updatedCollection.Entities.Add(
                        new Entity(entity.LogicalName)
                        {
                            Id = entity.Id,
                            [columnName] = current + 1  // Shift up by 1
                        }
                    );
                }

                return updatedCollection;
            }

            // UPDATE scenario: determine shift direction
            int direction;
            if (newOrder < oldOrder.Value)
                direction = +1; // Record moved up, so shift others down (increase their order)
            else if (newOrder > oldOrder.Value)
                direction = -1; // Record moved down, so shift others up (decrease their order)
            else
                return updatedCollection; // No movement, nothing to update

            // Apply the shift to all affected records
            foreach (var entity in collection.Entities)
            {
                var current = entity.GetAttributeValue<int?>(columnName) ?? 0;

                // Create a new entity with just the ID and updated order value
                updatedCollection.Entities.Add(
                    new Entity(entity.LogicalName)
                    {
                        Id = entity.Id,
                        [columnName] = current + direction  // Apply the shift
                    }
                );
            }

            return updatedCollection;
        }
    }
}
