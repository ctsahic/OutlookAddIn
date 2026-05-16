using System;
using System.Collections.Generic;

namespace OutlookAddIn.Models
{
    /// <summary>
    /// Defines configuration for a field including whether it's mandatory,
    /// if it should ignore validation errors, and its associated configuration key.
    /// </summary>
    public class FieldConfiguration
    {
        public string FieldName { get; set; }
        public bool IsMandatory { get; set; }
        public bool IgnoreInvalidValue { get; set; }
        public string ConfigurationKey { get; set; }
        public string DisplayName { get; set; }
        public string DefaultValue { get; set; }

        public FieldConfiguration()
        {
        }

        public FieldConfiguration(string fieldName, string configKey, bool isMandatory = true, bool ignoreInvalid = false)
        {
            FieldName = fieldName;
            ConfigurationKey = configKey;
            IsMandatory = isMandatory;
            IgnoreInvalidValue = ignoreInvalid;
            DisplayName = fieldName;
        }
    }
}
