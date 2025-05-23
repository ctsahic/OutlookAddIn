using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    partial class MyRibbon
    {
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.organizationUrlEditBox = this.Factory.CreateRibbonEditBox();
            this.projectNameEditBox = this.Factory.CreateRibbonEditBox();
            this.defaultAssigneeEditBox = this.Factory.CreateRibbonEditBox();
            this.btnSaveConfig = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.patEditBox = this.Factory.CreateRibbonEditBox();
            this.btnSavePat = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.tzahiButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Azure DevOps";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.organizationUrlEditBox);
            this.group1.Items.Add(this.projectNameEditBox);
            this.group1.Items.Add(this.defaultAssigneeEditBox);
            this.group1.Items.Add(this.btnSaveConfig);
            this.group1.Label = "Configuration";
            this.group1.Name = "group1";
            // 
            // organizationUrlEditBox
            // 
            this.organizationUrlEditBox.Label = "Organization URL";
            this.organizationUrlEditBox.Name = "organizationUrlEditBox";
            this.organizationUrlEditBox.SizeString = "https://dev.azure.com/your-org";
            this.organizationUrlEditBox.Text = null;
            // 
            // projectNameEditBox
            // 
            this.projectNameEditBox.Label = "Project Name";
            this.projectNameEditBox.Name = "projectNameEditBox";
            this.projectNameEditBox.SizeString = "your-project";
            this.projectNameEditBox.Text = null;
            // 
            // defaultAssigneeEditBox
            // 
            this.defaultAssigneeEditBox.Label = "Default Assignee";
            this.defaultAssigneeEditBox.Name = "defaultAssigneeEditBox";
            this.defaultAssigneeEditBox.SizeString = "Your Name";
            this.defaultAssigneeEditBox.Text = null;
            // 
            // btnSaveConfig
            // 
            this.btnSaveConfig.Label = "Save Config";
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveConfig_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.patEditBox);
            this.group2.Items.Add(this.btnSavePat);
            this.group2.Label = "Authentication";
            this.group2.Name = "group2";
            // 
            // patEditBox
            // 
            this.patEditBox.Label = "Personal Access Token";
            this.patEditBox.Name = "patEditBox";
            this.patEditBox.SizeString = "●●●●●●●●●●●●●●●●●●●●";
            this.patEditBox.Text = null;
            // 
            // btnSavePat
            // 
            this.btnSavePat.Label = "Save PAT";
            this.btnSavePat.Name = "btnSavePat";
            this.btnSavePat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSavePat_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.tzahiButton);
            this.group3.Label = "Actions";
            this.group3.Name = "group3";
            // 
            // tzahiButton
            // 
            this.tzahiButton.Label = "Create Bug";
            this.tzahiButton.Name = "tzahiButton";
            this.tzahiButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createBug_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox organizationUrlEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox projectNameEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox defaultAssigneeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox patEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSavePat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tzahiButton;
    }
}
