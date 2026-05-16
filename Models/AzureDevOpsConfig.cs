using System.Collections.Generic;

namespace OutlookAddIn.Models
{
    public class AzureDevOpsConfig
    {
        public string OrganizationUrl { get; set; }
        public string ProjectName { get; set; }
        public string DefaultAssignee { get; set; }
        
        /// <summary>
        /// Dynamic key-value configuration that can be used to populate fields
        /// and store custom parameters for work item creation.
        /// </summary>
        public Dictionary<string, string> DynamicParameters { get; set; } = new Dictionary<string, string>();
    }
}