using OutlookAddIn1.Models;

namespace OutlookAddIn1.Services
{
    public interface ICredentialService
    {
        string GetPat();
        void SavePat(string pat);
        AzureDevOpsConfig GetAzureDevOpsConfig();
        void SaveAzureDevOpsConfig(AzureDevOpsConfig config);
    }
}