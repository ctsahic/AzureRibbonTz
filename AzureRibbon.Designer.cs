using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AzureRibbonTz
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupUrl = this.Factory.CreateRibbonGroup();
            this.organizationUrlEditBox = this.Factory.CreateRibbonEditBox();
            this.projectNameEditBox = this.Factory.CreateRibbonEditBox();
            this.defaultAssigneeEditBox = this.Factory.CreateRibbonEditBox();
            this.groupPat = this.Factory.CreateRibbonGroup();
            this.patEditBox = this.Factory.CreateRibbonEditBox();
            this.btnSaveAll = this.Factory.CreateRibbonButton();
            this.groupActions = this.Factory.CreateRibbonGroup();
            this.tzahiButton = this.Factory.CreateRibbonButton();
            this.createStoryButton = this.Factory.CreateRibbonButton();
            this.updateItemButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupUrl.SuspendLayout();
            this.groupPat.SuspendLayout();
            this.groupActions.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupUrl);
            this.tab1.Groups.Add(this.groupPat);
            this.tab1.Groups.Add(this.groupActions);
            this.tab1.Label = "AzureTzaRibbon";
            this.tab1.Name = "tab1";
            // 
            // groupUrl
            // 
            this.groupUrl.Items.Add(this.organizationUrlEditBox);
            this.groupUrl.Items.Add(this.projectNameEditBox);
            this.groupUrl.Items.Add(this.defaultAssigneeEditBox);
            this.groupUrl.Label = "URL Configuration";
            this.groupUrl.Name = "groupUrl";
            // 
            // organizationUrlEditBox
            // 
            this.organizationUrlEditBox.Label = "URL";
            this.organizationUrlEditBox.Name = "organizationUrlEditBox";
            this.organizationUrlEditBox.SizeString = "XXXXXXXXXX";
            this.organizationUrlEditBox.Text = null;
            // 
            // projectNameEditBox
            // 
            this.projectNameEditBox.Label = "Project";
            this.projectNameEditBox.Name = "projectNameEditBox";
            this.projectNameEditBox.SizeString = "XXXXXXXXXX";
            this.projectNameEditBox.Text = null;
            // 
            // defaultAssigneeEditBox
            // 
            this.defaultAssigneeEditBox.Label = "Assignee";
            this.defaultAssigneeEditBox.Name = "defaultAssigneeEditBox";
            this.defaultAssigneeEditBox.SizeString = "XXXXXXXXXX";
            this.defaultAssigneeEditBox.Text = null;
            // 
            // groupPat
            // 
            this.groupPat.Items.Add(this.patEditBox);
            this.groupPat.Items.Add(this.btnSaveAll);
            this.groupPat.Label = "Authentication";
            this.groupPat.Name = "groupPat";
            // 
            // patEditBox
            // 
            this.patEditBox.Label = "PAT";
            this.patEditBox.Name = "patEditBox";
            this.patEditBox.SizeString = "XXXXXXXXXX";
            this.patEditBox.Text = null;
            // 
            // btnSaveAll
            // 
            this.btnSaveAll.Image = global::AzureRibbonTz.Properties.Resources.save_24;
            this.btnSaveAll.Label = "Save All";
            this.btnSaveAll.Name = "btnSaveAll";
            this.btnSaveAll.ShowImage = true;
            this.btnSaveAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAll_Click);
            // 
            // groupActions
            // 
            this.groupActions.Items.Add(this.tzahiButton);
            this.groupActions.Items.Add(this.createStoryButton);
            this.groupActions.Items.Add(this.updateItemButton);
            this.groupActions.Label = "Actions";
            this.groupActions.Name = "groupActions";
            // 
            // tzahiButton
            // 
            this.tzahiButton.Image = global::AzureRibbonTz.Properties.Resources.createBug;
            this.tzahiButton.Label = "Create Bug";
            this.tzahiButton.Name = "tzahiButton";
            this.tzahiButton.ShowImage = true;
            this.tzahiButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createBug_Click);
            // 
            // createStoryButton
            // 
            this.createStoryButton.Image = global::AzureRibbonTz.Properties.Resources.story;
            this.createStoryButton.Label = "Create Story";
            this.createStoryButton.Name = "createStoryButton";
            this.createStoryButton.ShowImage = true;
            this.createStoryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createStory_Click);
            // 
            // updateItemButton
            // 
            this.updateItemButton.Image = global::AzureRibbonTz.Properties.Resources.update_24;
            this.updateItemButton.Label = "Update Item";
            this.updateItemButton.Name = "updateItemButton";
            this.updateItemButton.ShowImage = true;
            this.updateItemButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateItem_Click);
            // 
            // MyRibbon
            // 
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
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPat;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox organizationUrlEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox projectNameEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox defaultAssigneeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox patEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tzahiButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createStoryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateItemButton;
    }
}