using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using OutlookAddIn.Models;

namespace OutlookAddIn.Services
{
    public class AzureDevOpsService : IAzureDevOpsService
    {
        private readonly AzureDevOpsConfig _config;

        public AzureDevOpsService(AzureDevOpsConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public async Task<WorkItem> CreateBugAsync(string title, string description, string pat)
        {
            return await CreateBugAsync(title, description, pat, new Dictionary<string, string>());
        }

        public async Task<WorkItem> CreateBugAsync(string title, string description, string pat, Dictionary<string, string> dynamicParameters)
        {
            if (string.IsNullOrWhiteSpace(title))
                throw new ArgumentNullException(nameof(title));
            if (string.IsNullOrWhiteSpace(pat))
                throw new ArgumentNullException(nameof(pat));

            var connection = VssConnectionHelper.CreateVssConnection(_config.OrganizationUrl, pat);

            var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            var patchDocument = new JsonPatchDocument
            {
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Title", Value = title },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/Microsoft.VSTS.TCM.ReproSteps", Value = description ?? string.Empty },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.History", Value = "Created from Outlook email" },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Tags", Value = "Created-From-Outlook" }
            };

            // Add assignee if configured
            if (!string.IsNullOrWhiteSpace(_config.DefaultAssignee))
            {
                patchDocument.Add(new JsonPatchOperation 
                { 
                    Operation = Operation.Add, 
                    Path = "/fields/System.AssignedTo", 
                    Value = _config.DefaultAssignee 
                });
            }

            // Add dynamic parameters as custom fields
            if (dynamicParameters != null && dynamicParameters.Count > 0)
            {
                foreach (var param in dynamicParameters)
                {
                    if (!string.IsNullOrWhiteSpace(param.Value))
                    {
                        // Map the parameter key to Azure DevOps field path
                        string fieldPath = MapParameterKeyToFieldPath(param.Key);
                        if (!string.IsNullOrEmpty(fieldPath))
                        {
                            patchDocument.Add(new JsonPatchOperation 
                            { 
                                Operation = Operation.Add, 
                                Path = fieldPath, 
                                Value = param.Value 
                            });
                        }
                    }
                }
            }

            return await witClient.CreateWorkItemAsync(patchDocument, _config.ProjectName, "Bug", bypassRules: true);
        }

        /// <summary>
        /// Maps a parameter key to an Azure DevOps field path.
        /// Supports multiple formats:
        /// - Simple field names: "Activity" -> "/fields/Microsoft.VSTS.Common.Activity"
        /// - Full reference names: "Microsoft.VSTS.CMMI.FoundInEnvironment" -> "/fields/Microsoft.VSTS.CMMI.FoundInEnvironment"
        /// - Field path format: "/fields/Microsoft.VSTS.CMMI.FoundInEnvironment" (returned as-is)
        /// </summary>
        private string MapParameterKeyToFieldPath(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return null;

            // If the key already starts with /fields/, return it as-is
            if (key.StartsWith("/fields/", StringComparison.OrdinalIgnoreCase))
                return key;

            // If key looks like a full reference name (contains dots and known namespaces), convert it to field path
            if (key.Contains(".") && (key.StartsWith("Microsoft.", StringComparison.OrdinalIgnoreCase) || 
                                      key.StartsWith("System.", StringComparison.OrdinalIgnoreCase) ||
                                      key.StartsWith("Custom.", StringComparison.OrdinalIgnoreCase)))
            {
                return $"/fields/{key}";
            }

            // Extract the last part of the key for simple name mapping
            var parts = key.Split('.');
            var fieldName = parts[parts.Length - 1];

            // Map common Azure DevOps field names to their full reference paths
            var fieldMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // System fields
                { "Title", "/fields/System.Title" },
                { "Description", "/fields/System.Description" },
                { "AssignedTo", "/fields/System.AssignedTo" },
                { "CreatedBy", "/fields/System.CreatedBy" },
                { "CreatedDate", "/fields/System.CreatedDate" },
                { "ChangedBy", "/fields/System.ChangedBy" },
                { "ChangedDate", "/fields/System.ChangedDate" },
                { "State", "/fields/System.State" },
                { "Reason", "/fields/System.Reason" },
                { "Area", "/fields/System.AreaPath" },
                { "AreaPath", "/fields/System.AreaPath" },
                { "AreaCode", "/fields/System.AreaPath" },
                { "Iteration", "/fields/System.IterationPath" },
                { "IterationPath", "/fields/System.IterationPath" },
                { "Tags", "/fields/System.Tags" },
                
                // Common VSTS fields
                { "Activity", "/fields/Microsoft.VSTS.Common.Activity" },
                { "ActivityGroup", "/fields/Microsoft.VSTS.Common.ActivityGroup" },
                { "Priority", "/fields/Microsoft.VSTS.Common.Priority" },
                { "Severity", "/fields/Microsoft.VSTS.Common.Severity" },
                { "BacklogPriority", "/fields/Microsoft.VSTS.Common.BacklogPriority" },
                { "BusinessValue", "/fields/Microsoft.VSTS.Common.BusinessValue" },
                { "StackRank", "/fields/Microsoft.VSTS.Common.StackRank" },
                
                // CMMI Process Template fields
                { "FoundInEnvironment", "/fields/Microsoft.VSTS.CMMI.FoundInEnvironment" },
                { "FoundIn", "/fields/Microsoft.VSTS.CMMI.FoundInEnvironment" },
                { "ResolvedInVersion", "/fields/Microsoft.VSTS.CMMI.ResolvedInVersion" },
                { "ResolvedIn", "/fields/Microsoft.VSTS.CMMI.ResolvedInVersion" },
                { "Blocked", "/fields/Microsoft.VSTS.CMMI.Blocked" },
                { "Issue", "/fields/Microsoft.VSTS.CMMI.Issue" },
                { "RequiresReview", "/fields/Microsoft.VSTS.CMMI.RequiresReview" },
                { "RequiresTest", "/fields/Microsoft.VSTS.CMMI.RequiresTest" },
                { "RootCause", "/fields/Microsoft.VSTS.CMMI.RootCause" },
                { "SystemInfo", "/fields/Microsoft.VSTS.CMMI.SystemInfo" },
                
                // TCM (Test Case Management) fields
                { "ReproSteps", "/fields/Microsoft.VSTS.TCM.ReproSteps" }
            };

            if (fieldMappings.ContainsKey(fieldName))
                return fieldMappings[fieldName];

            // If not a known field and contains dots, return as-is with /fields/ prefix
            if (key.Contains("."))
                return $"/fields/{key}";
            
            // Otherwise, treat as custom field
            return $"/fields/Custom.{fieldName}";
        }
    }
}