namespace OutlookAddIn1.Services
{
    public interface ICredentialService
    {
        void SavePat(string pat);
        string GetPat();
    }
}