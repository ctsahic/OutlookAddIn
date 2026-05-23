using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn.Models;

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

        public List<EmailAttachment> ProcessAttachments(Outlook.MailItem mail)
        {
            var results = new List<EmailAttachment>();
            if (mail?.Attachments == null) return results;

            string tempDir = Path.Combine(Path.GetTempPath(), "OutlookAddIn", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            foreach (Outlook.Attachment attachment in mail.Attachments)
            {
                try
                {
                    string safeFileName = string.Join("_", attachment.FileName.Split(Path.GetInvalidFileNameChars()));
                    string filePath = Path.Combine(tempDir, safeFileName);
                    attachment.SaveAsFile(filePath);

                    // Determine if inline by checking Content-ID
                    string contentId = null;
                    try
                    {
                        const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";
                        contentId = attachment.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID) as string;
                    }
                    catch { /* Property may not exist */ }

                    bool isInline = !string.IsNullOrEmpty(contentId);

                    results.Add(new EmailAttachment
                    {
                        FileName = safeFileName,
                        FilePath = filePath,
                        ContentId = contentId,
                        IsInline = isInline
                    });
                }
                catch { /* Skip specific attachments that fail */ }
            }
            return results;
        }
    }
}