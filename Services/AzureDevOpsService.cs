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
        /// Handles both standard field names and dot-notation for nested fields.
        /// Examples:
        ///   "Activity" -> "/fields/Microsoft.VSTS.Common.Activity"
        ///   "ActivityGroup.Activity" -> "/fields/Microsoft.VSTS.Common.ActivityGroup" and "/fields/Microsoft.VSTS.Common.Activity"
        /// </summary>
        private string MapParameterKeyToFieldPath(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return null;

            // Remove any dot notation and use the last part
            var parts = key.Split('.');
            var fieldName = parts[parts.Length - 1];

            // Map common Azure DevOps field names
            var fieldMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Activity", "/fields/Microsoft.VSTS.Common.Activity" },
                { "ActivityGroup", "/fields/Microsoft.VSTS.Common.ActivityGroup" },
                { "Area", "/fields/System.AreaPath" },
                { "AreaCode", "/fields/System.AreaPath" },
                { "Iteration", "/fields/System.IterationPath" },
                { "Priority", "/fields/Microsoft.VSTS.Common.Priority" },
                { "Severity", "/fields/Microsoft.VSTS.Common.Severity" },
                { "State", "/fields/System.State" },
                { "Tags", "/fields/System.Tags" },
                { "Reason", "/fields/System.Reason" }
            };

            if (fieldMappings.ContainsKey(fieldName))
                return fieldMappings[fieldName];

            // If not a known field, attempt to construct the path
            // This allows for custom fields with names like "Custom.MyField"
            return $"/fields/Custom.{fieldName}";
        }
    }
}