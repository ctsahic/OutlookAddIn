using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.Services
{
    public class OutlookEmailService : IEmailService
    {
        public string CleanDescription(string emailBody)
        {
            if (string.IsNullOrEmpty(emailBody))
                return string.Empty;

            // Remove HTML tags
            emailBody = Regex.Replace(emailBody, "<[^>]*>", string.Empty);
            // Remove extra whitespace
            emailBody = Regex.Replace(emailBody, @"\s+", " ");
            return emailBody.Trim();
        }

        public Outlook.MailItem GetSelectedEmail()
        {
            try
            {
                var app = new Outlook.Application();
                var explorer = app.ActiveExplorer();
                
                if (explorer?.Selection?.Count > 0)
                {
                    return explorer.Selection[1] as Outlook.MailItem;
                }
            }
            catch
            {
                // Return null if no valid email selected
            }
            
            return null;
        }
    }
}