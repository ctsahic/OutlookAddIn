using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace OutlookAddIn.Forms
{
    public class WorkItemCreatedDialog : Form
    {
        private string _workItemUrl;

        public WorkItemCreatedDialog(int workItemId, string organizationUrl, string projectName, string workItemType)
        {
            _workItemUrl = BuildWorkItemUrl(organizationUrl, projectName, workItemId);
            InitializeControls(workItemId, workItemType);
        }

        private string BuildWorkItemUrl(string organizationUrl, string projectName, int workItemId)
        {
            // Normalize URL
            organizationUrl = organizationUrl?.TrimEnd('/') ?? string.Empty;
            
            // Format: https://dev.azure.com/organization/project/_workitems/edit/123
            return $"{organizationUrl}/{projectName}/_workitems/edit/{workItemId}";
        }

        private void InitializeControls(int workItemId, string workItemType)
        {
            this.Text = "Work Item Created";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Width = 450;
            this.Height = 220;

            // Title Label
            var titleLabel = new Label
            {
                Text = $"{workItemType} created successfully!",
                Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(400, 30),
                AutoSize = false
            };

            // Item ID Label
            var idLabel = new Label
            {
                Text = $"Item ID: {workItemId}",
                Font = new System.Drawing.Font("Segoe UI", 10),
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(400, 25),
                AutoSize = false
            };

            // Link Label (clickable)
            var linkLabel = new LinkLabel
            {
                Text = "Open work item in Azure DevOps",
                Location = new System.Drawing.Point(20, 90),
                Size = new System.Drawing.Size(400, 25),
                AutoSize = false,
                LinkColor = System.Drawing.Color.Blue,
                VisitedLinkColor = System.Drawing.Color.Purple
            };
            linkLabel.Links.Add(0, linkLabel.Text.Length, _workItemUrl);
            linkLabel.LinkClicked += (s, e) => OpenLink(e.Link.LinkData.ToString());

            // OK Button
            var okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(350, 130),
                Size = new System.Drawing.Size(80, 30)
            };

            // Add controls to form
            this.Controls.Add(titleLabel);
            this.Controls.Add(idLabel);
            this.Controls.Add(linkLabel);
            this.Controls.Add(okButton);

            this.AcceptButton = okButton;
        }

        private void OpenLink(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open link: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
