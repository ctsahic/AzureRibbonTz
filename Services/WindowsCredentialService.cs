using CredentialManagement;
using OutlookAddIn1.Models;
using System;
using System.Text.Json;

namespace OutlookAddIn1.Services
{
    public class WindowsCredentialService : ICredentialService
    {
        private const string PAT_TARGET = "OutlookAddIn_PAT";
        private const string CONFIG_TARGET = "OutlookAddIn_AzureDevOpsConfig";

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
    }
}