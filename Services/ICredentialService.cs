using OutlookAddIn.Models;
using System.Collections.Generic;

namespace OutlookAddIn.Services
{
    public interface ICredentialService
    {
        string GetPat();
        void SavePat(string pat);
        AzureDevOpsConfig GetAzureDevOpsConfig();
        void SaveAzureDevOpsConfig(AzureDevOpsConfig config);
        
        /// <summary>
        /// Get field configuration that defines mandatory fields and their mappings.
        /// </summary>
        List<FieldConfiguration> GetFieldConfiguration();
        
        /// <summary>
        /// Save field configuration.
        /// </summary>
        void SaveFieldConfiguration(List<FieldConfiguration> configuration);
        
        /// <summary>
        /// Get a dynamic parameter value from configuration by key.
        /// </summary>
        string GetDynamicParameter(string key);
        
        /// <summary>
        /// Get all dynamic parameters.
        /// </summary>
        Dictionary<string, string> GetAllDynamicParameters();
        
        /// <summary>
        /// Save dynamic parameter (key-value pair).
        /// </summary>
        void SaveDynamicParameter(string key, string value);

        /// <summary>
        /// Delete a dynamic parameter by key.
        /// </summary>
        void DeleteDynamicParameter(string key);

        /// <summary>
        /// Clear all dynamic parameters.
        /// </summary>
        void ClearAllDynamicParameters();
    }
}