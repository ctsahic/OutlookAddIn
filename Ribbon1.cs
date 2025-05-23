using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.WebApi.Patch;

namespace OutlookAddIn1
{
    public partial class MyRibbon : RibbonBase
    {
        private readonly string _azureDevOpsUrl = "https://dev.azure.com/cohentsahi";
        private readonly string _projectName = "cohentzahi_agile";
        private readonly string _pat = "4RjWCK1VekyHkxbbnvIoiXpoad5yRkIXEiPtGtwsDIvvZvbMt36ZJQQJ99BEACAAAAAAAAAAAAASAZDO4KFM";

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private string CleanDescription(string emailBody)
        {
            // Remove any HTML tags if present
            emailBody = System.Text.RegularExpressions.Regex.Replace(emailBody, "<[^>]*>", string.Empty);

            // Remove excess whitespace
            emailBody = System.Text.RegularExpressions.Regex.Replace(emailBody, @"\s+", " ");

            // Trim the result
            return emailBody.Trim();
        }

        private async void tzahiButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mail = app.ActiveExplorer().Selection[1] as Outlook.MailItem;
                if (mail == null)
                {
                    System.Windows.Forms.MessageBox.Show("Please select a mail item.");
                    return;
                }

                string title = mail.Subject;
                string description = CleanDescription(mail.Body);

                VssConnection connection = new VssConnection(new Uri(_azureDevOpsUrl),
                    new VssBasicCredential(string.Empty, _pat));
                WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();

                var patchDocument = new JsonPatchDocument
                {
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Title",
                        Value = title
                    },
                    
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/Microsoft.VSTS.TCM.ReproSteps",
                        Value = description
                    },
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.History",
                        Value = "Created from Outlook email"
                    },
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.State",
                        Value = "New"
                    },
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.AssignedTo",
                        Value = "Tsahi Cohen" 
                    },
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/System.Tags",
                        Value = "Created-From-Outlook"
                    }
                };

                WorkItem result = await witClient.CreateWorkItemAsync(patchDocument, _projectName, "Bug");

                // Create a clickable link to the bug
                string bugUrl = $"{_azureDevOpsUrl}/{_projectName}/_workitems/edit/{result.Id}";
                System.Windows.Forms.MessageBox.Show(
                    $"Bug created successfully!\n\n Item ID: {result.Id}",
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
    }
}