using CredentialManagement;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn.Models;
using OutlookAddIn.Services;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class MyRibbon : RibbonBase
    {
        private ICredentialService _credentialService;
        private IAzureDevOpsService _azureDevOpsService;
        private IEmailService _emailService;
        private AzureDevOpsConfig _config;
        private bool _initialized;

        public MyRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (_initialized)
                return;

            _initialized = true;

            try
            {
                _credentialService = new WindowsCredentialService();
                _emailService = new OutlookEmailService();
                LoadSavedConfiguration();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during ribbon load: {ex.Message}");
            }
        }

        private void LoadSavedConfiguration()
        {
            if (_credentialService == null)
                return;

            try
            {
                var storedPat = _credentialService.GetPat();
                if (!string.IsNullOrEmpty(storedPat))
                    patEditBox.Text = new string('●', 8);

                var savedConfig = _credentialService.GetAzureDevOpsConfig();
                if (savedConfig != null)
                {
                    organizationUrlEditBox.Text = savedConfig.OrganizationUrl ?? string.Empty;
                    projectNameEditBox.Text = savedConfig.ProjectName ?? string.Empty;
                    defaultAssigneeEditBox.Text = savedConfig.DefaultAssignee ?? string.Empty;
                    
                    if (!string.IsNullOrEmpty(savedConfig.OrganizationUrl))
                    {
                        _config = savedConfig;
                        _azureDevOpsService = new AzureDevOpsService(_config);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading configuration: {ex.Message}");
            }
        }

        private bool ValidateAndSaveConfiguration()
        {
            try
            {
                var url = organizationUrlEditBox.Text?.Trim();
                var project = projectNameEditBox.Text?.Trim();
                var assignee = defaultAssigneeEditBox.Text?.Trim();
                var pat = patEditBox.Text?.Trim();

                if (string.IsNullOrWhiteSpace(url))
                {
                    MessageBox.Show("Please enter the Azure DevOps Organization URL.", "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(project))
                {
                    MessageBox.Show("Please enter the Project Name.", "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(assignee))
                {
                    MessageBox.Show("Please enter the Default Assignee.", "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(pat) || pat.Contains("●"))
                {
                    MessageBox.Show("Please enter a valid Personal Access Token (PAT).", "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                    url = "https://" + url;

                var uriTest = new Uri(url);
                
                _config = new AzureDevOpsConfig
                {
                    OrganizationUrl = url,
                    ProjectName = project,
                    DefaultAssignee = assignee
                };

                _credentialService.SaveAzureDevOpsConfig(_config);
                _credentialService.SavePat(pat);
                _azureDevOpsService = new AzureDevOpsService(_config);
                patEditBox.Text = new string('●', 8);
                
                return true;
            }
            catch (UriFormatException)
            {
                MessageBox.Show("Please enter a valid URL. Example: https://dev.azure.com/yourorganization",
                    "Invalid URL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving configuration: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private async void createBug_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_credentialService == null || _emailService == null)
                {
                    MessageBox.Show("Add-in failed to initialize properly. Please restart Outlook.", "Initialization Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!ValidateAndSaveConfiguration())
                    return;

                string pat = _credentialService.GetPat();
                if (string.IsNullOrEmpty(pat))
                {
                    MessageBox.Show("Please enter a valid Personal Access Token (PAT).", "Missing PAT",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var mail = _emailService.GetSelectedEmail();
                if (mail == null)
                {
                    MessageBox.Show("Please select a mail item first.", "No Selection",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var result = await _azureDevOpsService.CreateBugAsync(mail.Subject, _emailService.CleanDescription(mail.Body), pat);

                MessageBox.Show(
                    $"Bug created successfully!\n\nItem ID: {result.Id}",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating bug:\n\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
