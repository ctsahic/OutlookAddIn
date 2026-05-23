using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupUrl = this.Factory.CreateRibbonGroup();
            this.groupPat = this.Factory.CreateRibbonGroup();
            this.groupActions = this.Factory.CreateRibbonGroup();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.organizationUrlEditBox = this.Factory.CreateRibbonEditBox();
            this.projectNameEditBox = this.Factory.CreateRibbonEditBox();
            this.defaultAssigneeEditBox = this.Factory.CreateRibbonEditBox();
            this.patEditBox = this.Factory.CreateRibbonEditBox();
            this.tzahiButton = this.Factory.CreateRibbonButton();
            this.createUserStoryButton = this.Factory.CreateRibbonButton();
            this.settingsButton = this.Factory.CreateRibbonButton();
            this.viewFieldsButton = this.Factory.CreateRibbonButton();
            
            this.tab1.SuspendLayout();
            this.groupUrl.SuspendLayout();
            this.groupPat.SuspendLayout();
            this.groupActions.SuspendLayout();
            this.groupSettings.SuspendLayout();
            this.SuspendLayout();
            
            // tab1
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupUrl);
            this.tab1.Groups.Add(this.groupPat);
            this.tab1.Groups.Add(this.groupActions);
            this.tab1.Groups.Add(this.groupSettings);
            this.tab1.Label = "Azure";
            this.tab1.Name = "tab1";
            
            // groupUrl (Configuration)
            this.groupUrl.Items.Add(this.organizationUrlEditBox);
            this.groupUrl.Items.Add(this.projectNameEditBox);
            this.groupUrl.Items.Add(this.defaultAssigneeEditBox);
            this.groupUrl.Label = "URL Configuration";
            this.groupUrl.Name = "groupUrl";
            
            // organizationUrlEditBox
            this.organizationUrlEditBox.Label = "URL";
            this.organizationUrlEditBox.Name = "organizationUrlEditBox";
            this.organizationUrlEditBox.SizeString = "XXXXXXXXXX";
            
            // projectNameEditBox
            this.projectNameEditBox.Label = "Project";
            this.projectNameEditBox.Name = "projectNameEditBox";
            this.projectNameEditBox.SizeString = "XXXXXXXXXX";
            
            // defaultAssigneeEditBox
            this.defaultAssigneeEditBox.Label = "Assignee";
            this.defaultAssigneeEditBox.Name = "defaultAssigneeEditBox";
            this.defaultAssigneeEditBox.SizeString = "XXXXXXXXXX";
            
            // groupPat (Authentication)
            this.groupPat.Items.Add(this.patEditBox);
            this.groupPat.Label = "Authentication";
            this.groupPat.Name = "groupPat";
            
            // patEditBox
            this.patEditBox.Label = "PAT";
            this.patEditBox.Name = "patEditBox";
            this.patEditBox.SizeString = "XXXXXXXXXX";
            
            // groupActions
            this.groupActions.Items.Add(this.tzahiButton);
            this.groupActions.Items.Add(this.createUserStoryButton);
            this.groupActions.Label = "Actions";
            this.groupActions.Name = "groupActions";
            
            // tzahiButton
            this.tzahiButton.Label = "Create Bug";
            this.tzahiButton.Name = "tzahiButton";
            this.tzahiButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createBug_Click);

            // createUserStoryButton
            this.createUserStoryButton.Label = "Create User Story";
            this.createUserStoryButton.Name = "createUserStoryButton";
            this.createUserStoryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createUserStory_Click);

            // groupSettings
            this.groupSettings.Items.Add(this.settingsButton);
            this.groupSettings.Items.Add(this.viewFieldsButton);
            this.groupSettings.Label = "Settings";
            this.groupSettings.Name = "groupSettings";
            
            // settingsButton
            this.settingsButton.Label = "Configuration";
            this.settingsButton.Name = "settingsButton";
            this.settingsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.settings_Click);

            // viewFieldsButton
            this.viewFieldsButton.Label = "View Fields";
            this.viewFieldsButton.Name = "viewFieldsButton";
            this.viewFieldsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.viewFields_Click);
            
            // MyRibbon
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupUrl.ResumeLayout(false);
            this.groupUrl.PerformLayout();
            this.groupPat.ResumeLayout(false);
            this.groupPat.PerformLayout();
            this.groupActions.ResumeLayout(false);
            this.groupActions.PerformLayout();
            this.groupSettings.ResumeLayout(false);
            this.groupSettings.PerformLayout();
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPat;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox organizationUrlEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox projectNameEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox defaultAssigneeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox patEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tzahiButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createUserStoryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton settingsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton viewFieldsButton;
    }
}