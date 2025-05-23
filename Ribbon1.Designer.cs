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
            this.tzahiButton = this.Factory.CreateRibbonButton();
            this.patEditBox = this.Factory.CreateRibbonEditBox();
            this.btnSavePat = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.tzahiButton);
            this.group1.Items.Add(this.patEditBox);
            this.group1.Items.Add(this.btnSavePat);
            this.group1.Label = "My Group";
            this.group1.Name = "group1";
            // 
            // tzahiButton
            // 
            this.tzahiButton.Label = "Create Bug";
            this.tzahiButton.Name = "tzahiButton";
            this.tzahiButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tzahiButton_Click);
            // 
            // patEditBox
            // 
            this.patEditBox.Label = "PAT";
            this.patEditBox.Name = "patEditBox";
            // 
            // btnSavePat
            // 
            this.btnSavePat.Label = "Save";
            this.btnSavePat.Name = "btnSavePat";
            this.btnSavePat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSavePat_Click);
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
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox patEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSavePat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tzahiButton;
    }
}
