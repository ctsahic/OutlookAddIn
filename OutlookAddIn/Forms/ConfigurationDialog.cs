using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using OutlookAddIn.Services;

namespace OutlookAddIn.Forms
{
    public class ConfigurationDialog : Form
    {
        private readonly ConfigurationService _configurationService;
        private readonly ICredentialService _credentialService;
        private DataGridView _dynamicParametersGrid;
        private DataGridView _fieldConfigurationGrid;
        private Button _addParameterButton;
        private Button _removeParameterButton;
        private Button _saveButton;
        private Button _cancelButton;
        private TabControl _tabControl;

        public ConfigurationDialog(ConfigurationService configurationService, ICredentialService credentialService)
        {
            _configurationService = configurationService;
            _credentialService = credentialService;
            InitializeForm();
            LoadConfigurations();
        }

        private void InitializeForm()
        {
            this.Text = "Azure DevOps Configuration";
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Width = 800;
            this.Height = 550;
            this.MinimumSize = new Size(700, 450);

            // Tab Control - leaves room for buttons at bottom
            _tabControl = new TabControl
            {
                Location = new Point(10, 10),
                Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 110),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            // Tab 1: Dynamic Parameters
            var parameterTab = new TabPage("Dynamic Parameters");
            parameterTab.Padding = new Padding(5);
            InitializeParameterTab(parameterTab);
            _tabControl.TabPages.Add(parameterTab);

            // Tab 2: Field Configuration
            var fieldConfigTab = new TabPage("Field Configuration");
            fieldConfigTab.Padding = new Padding(5);
            InitializeFieldConfigurationTab(fieldConfigTab);
            _tabControl.TabPages.Add(fieldConfigTab);

            // Save Button
            _saveButton = new Button
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Location = new Point(this.ClientSize.Width - 180, this.ClientSize.Height - 50),
                Size = new Size(80, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _saveButton.Click += SaveButton_Click;

            // Cancel Button
            _cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(this.ClientSize.Width - 90, this.ClientSize.Height - 50),
                Size = new Size(80, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            this.Controls.Add(_tabControl);
            this.Controls.Add(_saveButton);
            this.Controls.Add(_cancelButton);

            this.AcceptButton = _saveButton;
            this.CancelButton = _cancelButton;
        }

        private void InitializeParameterTab(TabPage tab)
        {
            // Create a panel to hold controls
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            // Label
            var label = new Label
            {
                Text = "Dynamic Key-Value Parameters:",
                Location = new Point(5, 5),
                Size = new Size(panel.Width - 10, 20),
                AutoSize = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // DataGridView
            _dynamicParametersGrid = new DataGridView
            {
                Location = new Point(5, 30),
                Size = new Size(panel.Width - 10, panel.Height - 90),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };

            _dynamicParametersGrid.Columns.Add("Key", "Key");
            _dynamicParametersGrid.Columns.Add("Value", "Value");
            _dynamicParametersGrid.Columns[0].Width = 150;
            _dynamicParametersGrid.Columns[1].Width = 300;

            // Add Button
            _addParameterButton = new Button
            {
                Text = "Add Parameter",
                Location = new Point(5, panel.Height - 50),
                Size = new Size(120, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _addParameterButton.Click += (s, e) =>
            {
                _dynamicParametersGrid.Rows.Add("", "");
                _dynamicParametersGrid.CurrentCell = _dynamicParametersGrid.Rows[_dynamicParametersGrid.Rows.Count - 1].Cells[0];
            };

            // Remove Button
            _removeParameterButton = new Button
            {
                Text = "Remove",
                Location = new Point(135, panel.Height - 50),
                Size = new Size(120, 35),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _removeParameterButton.Click += (s, e) =>
            {
                if (_dynamicParametersGrid.SelectedRows.Count > 0)
                {
                    _dynamicParametersGrid.Rows.RemoveAt(_dynamicParametersGrid.SelectedRows[0].Index);
                }
            };

            panel.Controls.Add(label);
            panel.Controls.Add(_dynamicParametersGrid);
            panel.Controls.Add(_addParameterButton);
            panel.Controls.Add(_removeParameterButton);

            tab.Controls.Add(panel);
        }

        private void InitializeFieldConfigurationTab(TabPage tab)
        {
            // Create a panel to hold controls
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            // Label
            var label = new Label
            {
                Text = "Field Configuration:",
                Location = new Point(5, 5),
                Size = new Size(panel.Width - 10, 20),
                AutoSize = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // DataGridView
            _fieldConfigurationGrid = new DataGridView
            {
                Location = new Point(5, 30),
                Size = new Size(panel.Width - 10, panel.Height - 50),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                ReadOnly = true
            };

            _fieldConfigurationGrid.Columns.Add("FieldName", "Field Name");
            _fieldConfigurationGrid.Columns.Add("DisplayName", "Display Name");
            _fieldConfigurationGrid.Columns.Add("ConfigKey", "Configuration Key");
            _fieldConfigurationGrid.Columns.Add("IsMandatory", "Mandatory");
            _fieldConfigurationGrid.Columns.Add("IgnoreInvalid", "Ignore Invalid");
            _fieldConfigurationGrid.Columns.Add("DefaultValue", "Default Value");

            _fieldConfigurationGrid.Columns[0].Width = 110;
            _fieldConfigurationGrid.Columns[1].Width = 120;
            _fieldConfigurationGrid.Columns[2].Width = 130;
            _fieldConfigurationGrid.Columns[3].Width = 80;
            _fieldConfigurationGrid.Columns[4].Width = 110;
            _fieldConfigurationGrid.Columns[5].Width = 120;

            panel.Controls.Add(label);
            panel.Controls.Add(_fieldConfigurationGrid);

            tab.Controls.Add(panel);
        }

        private void LoadConfigurations()
        {
            // Load dynamic parameters
            var parameters = _configurationService.GetAllDynamicParameters();
            foreach (var param in parameters)
            {
                _dynamicParametersGrid.Rows.Add(param.Key, param.Value);
            }

            // Load field configurations
            var fieldConfigs = _configurationService.GetAllFieldConfigurations();
            foreach (var config in fieldConfigs)
            {
                _fieldConfigurationGrid.Rows.Add(
                    config.FieldName,
                    config.DisplayName,
                    config.ConfigurationKey,
                    config.IsMandatory,
                    config.IgnoreInvalidValue,
                    config.DefaultValue ?? string.Empty
                );
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Clear all existing dynamic parameters first
                _configurationService.GetAllDynamicParameters();
                _credentialService.ClearAllDynamicParameters();

                // Save only the rows that are currently in the grid
                foreach (DataGridViewRow row in _dynamicParametersGrid.Rows)
                {
                    var key = row.Cells[0].Value?.ToString();
                    var value = row.Cells[1].Value?.ToString();

                    // Only save non-empty keys
                    if (!string.IsNullOrWhiteSpace(key))
                    {
                        _credentialService.SaveDynamicParameter(key, value ?? string.Empty);
                    }
                }

                MessageBox.Show("Configuration saved successfully.", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving configuration: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
