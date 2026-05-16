using CredentialManagement;
using OutlookAddIn.Models;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace OutlookAddIn.Services
{
    public class WindowsCredentialService : ICredentialService
    {
        private const string PAT_TARGET = "OutlookAddIn_PAT";
        private const string CONFIG_TARGET = "OutlookAddIn_AzureDevOpsConfig";
        private const string FIELD_CONFIG_TARGET = "OutlookAddIn_FieldConfiguration";
        private const string DYNAMIC_PARAMS_TARGET = "OutlookAddIn_DynamicParameters";

        public string GetPat()
        {
            using (var cred = new Credential { Target = PAT_TARGET })
            {
                if (cred.Load())
                {
                    return cred.Password;
                }
                return null;
            }
        }

        public void SavePat(string pat)
        {
            using (var cred = new Credential
            {
                Target = PAT_TARGET,
                Password = pat,
                
            })
            {
                cred.Save();
            }
        }

        public AzureDevOpsConfig GetAzureDevOpsConfig()
        {
            using (var cred = new Credential { Target = CONFIG_TARGET })
            {
                if (cred.Load())
                {
                    try
                    {
                        return JsonSerializer.Deserialize<AzureDevOpsConfig>(cred.Password);
                    }
                    catch
                    {
                        return null;
                    }
                }
                return null;
            }
        }

        public void SaveAzureDevOpsConfig(AzureDevOpsConfig config)
        {
            var jsonConfig = JsonSerializer.Serialize(config);
            using (var cred = new Credential
            {
                Target = CONFIG_TARGET,
                Password = jsonConfig,
                
            })
            {
                cred.Save();
            }
        }

        public List<FieldConfiguration> GetFieldConfiguration()
        {
            using (var cred = new Credential { Target = FIELD_CONFIG_TARGET })
            {
                if (cred.Load())
                {
                    try
                    {
                        return JsonSerializer.Deserialize<List<FieldConfiguration>>(cred.Password) ?? new List<FieldConfiguration>();
                    }
                    catch
                    {
                        return new List<FieldConfiguration>();
                    }
                }
                return new List<FieldConfiguration>();
            }
        }

        public void SaveFieldConfiguration(List<FieldConfiguration> configuration)
        {
            var jsonConfig = JsonSerializer.Serialize(configuration);
            using (var cred = new Credential
            {
                Target = FIELD_CONFIG_TARGET,
                Password = jsonConfig,
            })
            {
                cred.Save();
            }
        }

        public string GetDynamicParameter(string key)
        {
            var parameters = GetAllDynamicParameters();
            return parameters.ContainsKey(key) ? parameters[key] : null;
        }

        public Dictionary<string, string> GetAllDynamicParameters()
        {
            using (var cred = new Credential { Target = DYNAMIC_PARAMS_TARGET })
            {
                if (cred.Load())
                {
                    try
                    {
                        return JsonSerializer.Deserialize<Dictionary<string, string>>(cred.Password) ?? new Dictionary<string, string>();
                    }
                    catch
                    {
                        return new Dictionary<string, string>();
                    }
                }
                return new Dictionary<string, string>();
            }
        }

        public void SaveDynamicParameter(string key, string value)
        {
            var parameters = GetAllDynamicParameters();
            parameters[key] = value;
            var jsonParams = JsonSerializer.Serialize(parameters);
            using (var cred = new Credential
            {
                Target = DYNAMIC_PARAMS_TARGET,
                Password = jsonParams,
            })
            {
                cred.Save();
            }
        }

        public void DeleteDynamicParameter(string key)
        {
            var parameters = GetAllDynamicParameters();
            if (parameters.ContainsKey(key))
            {
                parameters.Remove(key);
                var jsonParams = JsonSerializer.Serialize(parameters);
                using (var cred = new Credential
                {
                    Target = DYNAMIC_PARAMS_TARGET,
                    Password = jsonParams,
                })
                {
                    cred.Save();
                }
            }
        }

        public void ClearAllDynamicParameters()
        {
            var emptyParams = new Dictionary<string, string>();
            var jsonParams = JsonSerializer.Serialize(emptyParams);
            using (var cred = new Credential
            {
                Target = DYNAMIC_PARAMS_TARGET,
                Password = jsonParams,
            })
            {
                cred.Save();
            }
        }
    }
}