using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace DcoumentRouterPlugins
{
    public class InitializeChildRecordOnCreate : PluginBase
    {
        // Entity references
        private const string DocumentRouter = "cr8d2_routingsummary";
        private const string SignatureFile = "cr8d2_documentroutersignaturefile";

        // Lookup on document router
        private const string SignatureFileLookup = "cr8d2_documentroutersignaturefile";

        public InitializeChildRecordOnCreate() : base(typeof(InitializeChildRecordOnCreate)) { }

        protected override void ExecuteCdsPlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var sysService = localPluginContext.SystemUserService;
            var tracer = localPluginContext.TracingService;

            if (context.MessageName != "Create" || context.Stage != 40)
                throw new InvalidPluginExecutionException("Invalid execution context. Plugin must be registered on Post-Operation Create.");

            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity target))
                throw new InvalidPluginExecutionException("Target entity not found in input parameters.");

            if (target.LogicalName != DocumentRouter) return;

            try
            {
                tracer.Trace("Starting creation of Document Router Signature File child record");

                Entity newSignatureRecord = new Entity(SignatureFile);

                // Give child record same name as document router parent
                string parentName = target.Contains("cr8d2_name") ? target["cr8d2_name"].ToString() : "Unknown Router";
                newSignatureRecord["cr8d2_name"] = $"Signature Payload - {parentName}";

                // Create child record
                Guid childRecordId = sysService.Create(newSignatureRecord);
                tracer.Trace($"Successfully created child signature file record with ID: {childRecordId}");

                // Set child lookup on document router
                Entity parentUpdate = new Entity(DocumentRouter)
                {
                    Id = target.Id
                };

                parentUpdate[SignatureFileLookup] = new EntityReference(SignatureFile, childRecordId);

                sysService.Update(parentUpdate);
                tracer.Trace("Successfully linked the new signature file record to parent");

            }
            catch (Exception ex)
            {

                tracer.Trace($"Error in InitializeChildRecordOnCreate : {ex.Message}");
                throw new InvalidPluginExecutionException($"An error occurred while generating the signature file record: {ex:Message}", ex);
            }






        }
    }

}
