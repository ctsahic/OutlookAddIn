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
using Microsoft.Office.Tools.Ribbon;
using OutlookAddIn1.Models;
using OutlookAddIn1.Services;
using System.Net;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class MyRibbon : RibbonBase
    {
        private readonly ICredentialService _credentialService;
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IEmailService _emailService;

        public MyRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            var config = new AzureDevOpsConfig
            {
                OrganizationUrl = "https://dev.azure.com/cohentsahi",
                ProjectName = "cohentzahi_agile",
                DefaultAssignee = "Tsahi Cohen"
            };

            _credentialService = new WindowsCredentialService();
            _azureDevOpsService = new AzureDevOpsService(config);
            _emailService = new OutlookEmailService();
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var storedPat = _credentialService.GetPat();
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
                _credentialService.SavePat(enteredPat);
                patEditBox.Text = new string('●', 8);
                MessageBox.Show("PAT saved successfully.");
            }
        }

        private async void tzahiButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string pat = _credentialService.GetPat();
                if (string.IsNullOrEmpty(pat))
                {
                    MessageBox.Show("Please enter and save your PAT first.");
                    return;
                }

                var mail = _emailService.GetSelectedEmail();
                if (mail == null)
                {
                    MessageBox.Show("Please select a mail item.");
                    return;
                }

                string title = mail.Subject;
                string description = _emailService.CleanDescription(mail.Body);

                var result = await _azureDevOpsService.CreateBugAsync(title, description, pat);

                MessageBox.Show(
                    $"Bug created successfully!\n\nItem ID: {result.Id}",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error creating bug:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
