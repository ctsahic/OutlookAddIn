using CredentialManagement;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn.Models;
using OutlookAddIn.Services;
using OutlookAddIn.Forms;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class MyRibbon : RibbonBase
    {
        private ICredentialService _credentialService;
        private ConfigurationService _configurationService;
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
                _configurationService = new ConfigurationService(_credentialService);
                _configurationService.InitializeDefaultConfigurations();
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

                // Try to populate from dynamic configuration if values are missing
                url = _configurationService.GetFieldValue("OrganizationUrl", url);
                project = _configurationService.GetFieldValue("ProjectName", project);
                assignee = _configurationService.GetFieldValue("DefaultAssignee", assignee);

                // Validate using configuration service
                var urlValidation = _configurationService.ValidateField("OrganizationUrl", url);
                if (!urlValidation.IsValid)
                {
                    MessageBox.Show(urlValidation.ErrorMessage, "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                var projectValidation = _configurationService.ValidateField("ProjectName", project);
                if (!projectValidation.IsValid)
                {
                    MessageBox.Show(projectValidation.ErrorMessage, "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                var assigneeValidation = _configurationService.ValidateField("DefaultAssignee", assignee);
                if (!assigneeValidation.IsValid)
                {
                    MessageBox.Show(assigneeValidation.ErrorMessage, "Validation Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Check if PAT is masked (user hasn't changed it from the saved value)
                if (pat.Contains("●"))
                {
                    // PAT is masked, try to use the previously saved PAT
                    pat = _credentialService.GetPat();
                    if (string.IsNullOrEmpty(pat))
                    {
                        MessageBox.Show("Please enter a valid Personal Access Token (PAT).", "Validation Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                }
                else if (string.IsNullOrWhiteSpace(pat))
                {
                    // PAT field is empty
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

                // Get dynamic parameters to include in the work item
                var dynamicParams = _configurationService.GetAllDynamicParameters();

                var result = await _azureDevOpsService.CreateBugAsync(mail.Subject, _emailService.CleanDescription(mail.Body), pat, dynamicParams);

                // Show custom dialog with link to the created work item
                using (var dialog = new WorkItemCreatedDialog(result.Id ?? 0, _config.OrganizationUrl, _config.ProjectName))
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating bug:\n\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void settings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var dialog = new ConfigurationDialog(_configurationService, _credentialService))
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening configuration: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
