using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using OutlookAddIn.Models;

namespace OutlookAddIn.Services
{
    /// <summary>
    /// Service to retrieve metadata information from Azure DevOps,
    /// including available fields, their types, and allowed values.
    /// </summary>
    public class AzureDevOpsMetadataService
    {
        private readonly AzureDevOpsConfig _config;

        public AzureDevOpsMetadataService(AzureDevOpsConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        /// <summary>
        /// Get all available fields for the Bug work item type.
        /// </summary>
        public async Task<List<WorkItemField>> GetBugFieldsAsync(string pat)
        {
            if (string.IsNullOrWhiteSpace(pat))
                throw new ArgumentNullException(nameof(pat));

            var connection = VssConnectionHelper.CreateVssConnection(_config.OrganizationUrl, pat);

            var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            try
            {
                // Get all fields
                var fields = await witClient.GetFieldsAsync();
                return fields.ToList();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to retrieve fields from Azure DevOps: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Get the Bug work item type with all its fields and allowed values.
        /// </summary>
        public async Task<WorkItemType> GetBugWorkItemTypeAsync(string pat)
        {
            if (string.IsNullOrWhiteSpace(pat))
                throw new ArgumentNullException(nameof(pat));

            var connection = VssConnectionHelper.CreateVssConnection(_config.OrganizationUrl, pat);

            var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            try
            {
                // Get the Bug work item type for the project
                var bugWorkItemType = await witClient.GetWorkItemTypeAsync(_config.ProjectName, "Bug");
                return bugWorkItemType;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to retrieve Bug work item type from Azure DevOps: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Get all fields for the Bug work item type with their metadata.
        /// Returns a formatted list of field information.
        /// </summary>
        public async Task<List<FieldMetadata>> GetBugFieldMetadataAsync(string pat)
        {
            if (string.IsNullOrWhiteSpace(pat))
                throw new ArgumentNullException(nameof(pat));

            var bugWorkItemType = await GetBugWorkItemTypeAsync(pat);
            var fields = await GetBugFieldsAsync(pat);

            var fieldMetadata = new List<FieldMetadata>();

            foreach (var fieldReference in bugWorkItemType.Fields)
            {
                var field = fields.FirstOrDefault(f => f.ReferenceName == fieldReference.ReferenceName);
                if (field != null)
                {
                    var metadata = new FieldMetadata
                    {
                        ReferenceName = field.ReferenceName,
                        Name = field.Name,
                        Type = field.Type.ToString(),
                        Usage = field.Usage.ToString(),
                        Description = field.Description ?? ""
                    };

                    fieldMetadata.Add(metadata);
                }
            }

            return fieldMetadata.OrderBy(f => f.Name).ToList();
        }

        /// <summary>
        /// Get a sample of a created work item to see all populated fields.
        /// </summary>
        public async Task<WorkItem> GetWorkItemAsync(int workItemId, string pat)
        {
            if (string.IsNullOrWhiteSpace(pat))
                throw new ArgumentNullException(nameof(pat));

            var connection = VssConnectionHelper.CreateVssConnection(_config.OrganizationUrl, pat);

            var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            try
            {
                // Get the work item with all fields
                var workItem = await witClient.GetWorkItemAsync(_config.ProjectName, workItemId, expand: WorkItemExpand.All);
                return workItem;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to retrieve work item {workItemId} from Azure DevOps: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Print all fields of a work item to debug output.
        /// </summary>
        public void PrintWorkItemFields(WorkItem workItem)
        {
            System.Diagnostics.Debug.WriteLine("========== WORK ITEM FIELDS ==========");
            System.Diagnostics.Debug.WriteLine($"Work Item ID: {workItem.Id}");
            System.Diagnostics.Debug.WriteLine($"Revision: {workItem.Rev}");
            System.Diagnostics.Debug.WriteLine("");
            System.Diagnostics.Debug.WriteLine("========== FIELD VALUES ==========");

            if (workItem.Fields != null && workItem.Fields.Count > 0)
            {
                foreach (var field in workItem.Fields.OrderBy(f => f.Key))
                {
                    var value = field.Value?.ToString() ?? "[NULL]";
                    System.Diagnostics.Debug.WriteLine($"{field.Key}: {value}");
                }
            }

            System.Diagnostics.Debug.WriteLine("========== END FIELDS ==========");
        }

        /// <summary>
        /// Get a formatted string of all work item fields for display purposes.
        /// </summary>
        public string GetWorkItemFieldsAsString(WorkItem workItem)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("========== WORK ITEM FIELDS ==========");
            sb.AppendLine($"Work Item ID: {workItem.Id}");
            sb.AppendLine($"Revision: {workItem.Rev}");
            sb.AppendLine("");
            sb.AppendLine("========== FIELD VALUES ==========");

            if (workItem.Fields != null && workItem.Fields.Count > 0)
            {
                foreach (var field in workItem.Fields.OrderBy(f => f.Key))
                {
                    var value = field.Value?.ToString() ?? "[NULL]";
                    sb.AppendLine($"{field.Key}: {value}");
                }
            }

            sb.AppendLine("========== END FIELDS ==========");
            return sb.ToString();
        }
    }

    /// <summary>
    /// Metadata about a work item field.
    /// </summary>
    public class FieldMetadata
    {
        public string ReferenceName { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Usage { get; set; }
        public string Description { get; set; }

        public override string ToString()
        {
            return $"[{Name}] {ReferenceName}\n" +
                   $"  Type: {Type}\n" +
                   $"  Usage: {Usage}\n" +
                   $"  Description: {Description}";
        }
    }
}
