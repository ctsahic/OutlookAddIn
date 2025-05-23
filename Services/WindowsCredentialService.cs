using CredentialManagement;

namespace OutlookAddIn1.Services
{
    public class WindowsCredentialService : ICredentialService
    {
        private const string CREDENTIAL_TARGET = "AzureDevOpsBugReporter";

        public string GetPat()
        {
            using (var credential = new Credential())
            {
                credential.Target = CREDENTIAL_TARGET;
                credential.Type = CredentialType.Generic;
                return credential.Load() ? credential.Password : null;
            }
        }

        public void SavePat(string pat)
        {
            using (var credential = new Credential())
            {
                credential.Target = CREDENTIAL_TARGET;
                credential.Username = "PAT";
                credential.Password = pat;
                credential.Type = CredentialType.Generic;
                credential.PersistanceType = PersistanceType.LocalComputer;
                credential.Save();
            }
        }
    }
}