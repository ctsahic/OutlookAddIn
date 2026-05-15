using OutlookAddIn.Models;

namespace OutlookAddIn.Services
{
    public interface ICredentialService
    {
        string GetPat();
        void SavePat(string pat);
        AzureDevOpsConfig GetAzureDevOpsConfig();
        void SaveAzureDevOpsConfig(AzureDevOpsConfig config);
    }
}