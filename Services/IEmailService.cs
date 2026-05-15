using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.Services
{
    public interface IEmailService
    {
        Outlook.MailItem GetSelectedEmail();
        string CleanDescription(string emailBody);
    }
}