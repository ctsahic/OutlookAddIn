using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using OutlookAddIn.Services;

namespace OutlookAddIn.Forms
{
    public class FieldMetadataDialog : Form
    {
        private DataGridView _fieldsGrid;
        private Button _closeButton;
        private Button _copyButton;
        private Label _statusLabel;
        private ProgressBar _progressBar;
        private List<FieldMetadata> _fieldMetadata;

        public FieldMetadataDialog()
        {
            InitializeForm();
        }

        private void InitializeForm()
        {
            this.Text = "Azure DevOps Bug Fields Metadata";
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Width = 1000;
            this.Height = 600;
            this.MinimumSize = new Size(800, 400);

            // Progress Bar
            _progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                Location = new Point(10, 10),
                Size = new Size(this.ClientSize.Width - 20, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false
            };

            // Status Label
            _statusLabel = new Label
            {
                Text = "Ready",
                Location = new Point(10, 35),
                Size = new Size(this.ClientSize.Width - 20, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // DataGridView
            _fieldsGrid = new DataGridView
            {
                Location = new Point(10, 60),
                Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 130),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToOrderColumns = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                ReadOnly = true,
                MultiSelect = true
            };

            // Initialize columns
            _fieldsGrid.Columns.Add("Name", "Field Name");
            _fieldsGrid.Columns.Add("ReferenceName", "Reference Name");
            _fieldsGrid.Columns.Add("Type", "Type");
            _fieldsGrid.Columns.Add("Usage", "Usage");
            _fieldsGrid.Columns.Add("Description", "Description");

            // Set column widths
            _fieldsGrid.Columns["Name"].Width = 150;
            _fieldsGrid.Columns["ReferenceName"].Width = 250;
            _fieldsGrid.Columns["Type"].Width = 80;
            _fieldsGrid.Columns["Usage"].Width = 80;
            _fieldsGrid.Columns["Description"].Width = 300;

            // Copy Button
            _copyButton = new Button
            {
                Text = "Copy Selected to Clipboard",
                Location = new Point(10, this.ClientSize.Height - 50),
                Size = new Size(150, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _copyButton.Click += (s, e) => CopySelectedFields();

            // Close Button
            _closeButton = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                Location = new Point(this.ClientSize.Width - 100, this.ClientSize.Height - 50),
                Size = new Size(90, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            this.Controls.Add(_progressBar);
            this.Controls.Add(_statusLabel);
            this.Controls.Add(_fieldsGrid);
            this.Controls.Add(_copyButton);
            this.Controls.Add(_closeButton);

            this.AcceptButton = _closeButton;
        }

        public async void LoadFieldsAsync(AzureDevOpsMetadataService metadataService, string pat)
        {
            try
            {
                _progressBar.Visible = true;
                _statusLabel.Text = "Loading fields metadata from Azure DevOps...";
                this.Refresh();

                _fieldMetadata = await metadataService.GetBugFieldMetadataAsync(pat);

                _fieldsGrid.Rows.Clear();
                foreach (var field in _fieldMetadata)
                {
                    _fieldsGrid.Rows.Add(
                        field.Name,
                        field.ReferenceName,
                        field.Type,
                        field.Usage,
                        field.Description
                    );
                }

                _statusLabel.Text = $"Loaded {_fieldMetadata.Count} fields";
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Error: {ex.Message}";
                MessageBox.Show($"Error loading fields: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _progressBar.Visible = false;
            }
        }

        public async void LoadWorkItemFieldsAsync(AzureDevOpsMetadataService metadataService, int workItemId, string pat)
        {
            try
            {
                _progressBar.Visible = true;
                _statusLabel.Text = $"Loading work item {workItemId} fields from Azure DevOps...";
                this.Refresh();

                var workItem = await metadataService.GetWorkItemAsync(workItemId, pat);
                metadataService.PrintWorkItemFields(workItem);

                _fieldsGrid.Rows.Clear();
                if (workItem.Fields != null)
                {
                    foreach (var field in workItem.Fields.OrderBy(f => f.Key))
                    {
                        var value = field.Value?.ToString() ?? "[NULL]";
                        _fieldsGrid.Rows.Add(
                            field.Key,
                            field.Key,
                            "[Value]",
                            "",
                            value
                        );
                    }
                }

                _statusLabel.Text = $"Loaded {workItem.Fields?.Count ?? 0} fields from work item {workItemId}";
                
                // Also print to debug output
                var fieldString = metadataService.GetWorkItemFieldsAsString(workItem);
                System.Diagnostics.Debug.WriteLine(fieldString);
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Error: {ex.Message}";
                MessageBox.Show($"Error loading work item fields: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _progressBar.Visible = false;
            }
        }

        private void CopySelectedFields()
        {
            if (_fieldsGrid.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select at least one field to copy.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var sb = new System.Text.StringBuilder();
            foreach (DataGridViewRow row in _fieldsGrid.SelectedRows)
            {
                sb.AppendLine($"Key: {row.Cells["ReferenceName"].Value}");
                sb.AppendLine($"Value: {row.Cells["Description"].Value}");
                sb.AppendLine();
            }

            try
            {
                System.Windows.Forms.Clipboard.SetText(sb.ToString());
                _statusLabel.Text = $"Copied {_fieldsGrid.SelectedRows.Count} field(s) to clipboard";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying to clipboard: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
