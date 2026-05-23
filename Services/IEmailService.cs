using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn.Models;
using System.Collections.Generic;

namespace OutlookAddIn.Services
{
    public interface IEmailService
    {
        Outlook.MailItem GetSelectedEmail();
        string CleanDescription(string emailBody);
        List<EmailAttachment> ProcessAttachments(Outlook.MailItem mail);
    }
}