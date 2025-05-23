using CredentialManagement;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using OutlookAddIn;
using System;
using static Microsoft.TeamFoundation.Common.Internal.NativeMethods;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public partial class MyRibbon : RibbonBase
    {
        private readonly string _azureDevOpsUrl = "https://dev.azure.com/cohentsahi";
        private readonly string _projectName = "cohentzahi_agile";

        public MyRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var storedPat = GetPat();
            if (!string.IsNullOrEmpty(storedPat))
            {
                patEditBox.Text = new string('●', 8);
            }
        }

        private void btnSavePat_Click(object sender, RibbonControlEventArgs e)
        {
            string enteredPat = patEditBox.Text;
            if (!string.IsNullOrWhiteSpace(enteredPat) && !enteredPat.Contains("●"))
            {
                SavePat(enteredPat);
                patEditBox.Text = new string('●', 8);
                System.Windows.Forms.MessageBox.Show("PAT saved successfully.");
            }
        }

        private string CleanDescription(string emailBody)
        {
            emailBody = System.Text.RegularExpressions.Regex.Replace(emailBody, "<[^>]*>", string.Empty);
            emailBody = System.Text.RegularExpressions.Regex.Replace(emailBody, @"\s+", " ");
            return emailBody.Trim();
        }

        private async void tzahiButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                string pat = GetPat();
                if (string.IsNullOrEmpty(pat))
                {
                    System.Windows.Forms.MessageBox.Show("Please enter and save your PAT first.");
                    return;
                }

                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mail = app.ActiveExplorer().Selection[1] as Outlook.MailItem;
                if (mail == null)
                {
                    System.Windows.Forms.MessageBox.Show("Please select a mail item.");
                    return;
                }

                string title = mail.Subject;
                string description = CleanDescription(mail.Body);

                var connection = new VssConnection(new Uri(_azureDevOpsUrl), new VssBasicCredential(string.Empty, pat));
                var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

                var patchDocument = new JsonPatchDocument
                {
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Title", Value = title },
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/Microsoft.VSTS.TCM.ReproSteps", Value = description },
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.History", Value = "Created from Outlook email" },
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.State", Value = "New" },
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.AssignedTo", Value = "Tsahi Cohen" },
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Tags", Value = "Created-From-Outlook" }
                };

                WorkItem result = await witClient.CreateWorkItemAsync(patchDocument, _projectName, "Bug");

                System.Windows.Forms.MessageBox.Show(
                    $"Bug created successfully!\n\nItem ID: {result.Id}",
                    "Success",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error creating bug:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void SavePat(string pat)
        {
            var cred = new Credential { Target = "AzureDevOpsBugReporter", Password = pat, PersistanceType = PersistanceType.LocalComputer };
            cred.Save();
        }

        private string GetPat()
        {
            var cred = new Credential { Target = "AzureDevOpsBugReporter" };
            return cred.Load() ? cred.Password : null;
        }
    }
}
