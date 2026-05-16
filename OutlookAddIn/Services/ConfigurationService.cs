using OutlookAddIn.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAddIn.Services
{
    /// <summary>
    /// Service to manage field configuration and validation with support for dynamic parameters.
    /// </summary>
    public class ConfigurationService
    {
        private readonly ICredentialService _credentialService;
        private List<FieldConfiguration> _fieldConfigurations;

        public ConfigurationService(ICredentialService credentialService)
        {
            _credentialService = credentialService;
            _fieldConfigurations = _credentialService.GetFieldConfiguration();
        }

        /// <summary>
        /// Initialize default field configurations if none exist.
        /// </summary>
        public void InitializeDefaultConfigurations()
        {
            if (_fieldConfigurations.Count == 0)
            {
                _fieldConfigurations = new List<FieldConfiguration>
                {
                    new FieldConfiguration("OrganizationUrl", "org_url", true, false) 
                    { 
                        DisplayName = "Organization URL" 
                    },
                    new FieldConfiguration("ProjectName", "project_name", true, false)
                    { 
                        DisplayName = "Project Name" 
                    },
                    new FieldConfiguration("DefaultAssignee", "default_assignee", true, false)
                    { 
                        DisplayName = "Default Assignee" 
                    },
                };

                SaveFieldConfigurations();
            }
        }

        /// <summary>
        /// Get field configuration by field name.
        /// </summary>
        public FieldConfiguration GetFieldConfiguration(string fieldName)
        {
            return _fieldConfigurations.FirstOrDefault(fc => fc.FieldName == fieldName);
        }

        /// <summary>
        /// Get all field configurations.
        /// </summary>
        public List<FieldConfiguration> GetAllFieldConfigurations()
        {
            return new List<FieldConfiguration>(_fieldConfigurations);
        }

        /// <summary>
        /// Add or update a field configuration.
        /// </summary>
        public void SetFieldConfiguration(FieldConfiguration config)
        {
            var existing = _fieldConfigurations.FirstOrDefault(fc => fc.FieldName == config.FieldName);
            if (existing != null)
            {
                _fieldConfigurations.Remove(existing);
            }
            _fieldConfigurations.Add(config);
            SaveFieldConfigurations();
        }

        /// <summary>
        /// Get value for a field, trying dynamic parameters if IgnoreInvalidValue is true.
        /// </summary>
        public string GetFieldValue(string fieldName, string currentValue)
        {
            var config = GetFieldConfiguration(fieldName);
            
            if (config == null)
                return currentValue;

            // If current value is valid, use it
            if (!string.IsNullOrWhiteSpace(currentValue))
                return currentValue;

            // If ignore invalid is enabled, try to get from dynamic parameters
            if (config.IgnoreInvalidValue && !string.IsNullOrEmpty(config.ConfigurationKey))
            {
                var dynamicValue = _credentialService.GetDynamicParameter(config.ConfigurationKey);
                if (!string.IsNullOrEmpty(dynamicValue))
                    return dynamicValue;
            }

            // Use default value if available
            if (!string.IsNullOrEmpty(config.DefaultValue))
                return config.DefaultValue;

            return currentValue;
        }

        /// <summary>
        /// Validate a field value according to its configuration.
        /// </summary>
        public ValidationResult ValidateField(string fieldName, string value)
        {
            var config = GetFieldConfiguration(fieldName);
            
            if (config == null)
                return ValidationResult.Success();

            // If mandatory and empty
            if (config.IsMandatory && string.IsNullOrWhiteSpace(value))
            {
                // Try to get from dynamic parameters
                if (!string.IsNullOrEmpty(config.ConfigurationKey))
                {
                    var dynamicValue = _credentialService.GetDynamicParameter(config.ConfigurationKey);
                    if (!string.IsNullOrEmpty(dynamicValue))
                        return ValidationResult.Success();
                }

                return ValidationResult.Failure($"{config.DisplayName} is required and not provided in configuration.");
            }

            // If ignore invalid is true, allow empty values
            if (config.IgnoreInvalidValue && string.IsNullOrWhiteSpace(value))
            {
                // But still try to populate from dynamic parameters
                if (!string.IsNullOrEmpty(config.ConfigurationKey))
                {
                    var dynamicValue = _credentialService.GetDynamicParameter(config.ConfigurationKey);
                    if (!string.IsNullOrEmpty(dynamicValue))
                        return ValidationResult.Success();
                }
                return ValidationResult.Success();
            }

            return ValidationResult.Success();
        }

        /// <summary>
        /// Save a dynamic parameter value.
        /// </summary>
        public void SetDynamicParameter(string key, string value)
        {
            _credentialService.SaveDynamicParameter(key, value);
        }

        /// <summary>
        /// Get a dynamic parameter value.
        /// </summary>
        public string GetDynamicParameter(string key)
        {
            return _credentialService.GetDynamicParameter(key);
        }

        /// <summary>
        /// Get all dynamic parameters.
        /// </summary>
        public Dictionary<string, string> GetAllDynamicParameters()
        {
            return _credentialService.GetAllDynamicParameters();
        }

        private void SaveFieldConfigurations()
        {
            _credentialService.SaveFieldConfiguration(_fieldConfigurations);
        }
    }

    /// <summary>
    /// Represents the result of a field validation.
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; }

        public static ValidationResult Success()
        {
            return new ValidationResult { IsValid = true };
        }

        public static ValidationResult Failure(string message)
        {
            return new ValidationResult { IsValid = false, ErrorMessage = message };
        }
    }
}
