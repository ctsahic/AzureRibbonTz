using CredentialManagement;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System;
using System.Configuration;
using static Microsoft.TeamFoundation.Common.Internal.NativeMethods;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn1.Models;
using OutlookAddIn1.Services;
using System.Net;
using System.Windows.Forms;

namespace AzureRibbonTz
{
    public partial class MyRibbon : RibbonBase
    {
        private readonly ICredentialService _credentialService;
        private IAzureDevOpsService _azureDevOpsService;  // Changed from IAsyncLazy<T>
        private readonly IEmailService _emailService;
        private AzureDevOpsConfig _config;

        public MyRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            _credentialService = new WindowsCredentialService();
            _emailService = new OutlookEmailService();
            _config = new AzureDevOpsConfig();
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var storedPat = _credentialService.GetPat();
            if (!string.IsNullOrEmpty(storedPat))
            {
                patEditBox.Text = new string('●', 8);
            }

            // Load saved configuration
            var savedConfig = _credentialService.GetAzureDevOpsConfig();
            if (savedConfig != null)
            {
                organizationUrlEditBox.Text = savedConfig.OrganizationUrl;
                projectNameEditBox.Text = savedConfig.ProjectName;
                defaultAssigneeEditBox.Text = savedConfig.DefaultAssignee;
                _config = savedConfig;
                _azureDevOpsService = new AzureDevOpsService(_config);
            }
        }

        private void btnSaveAll_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(organizationUrlEditBox.Text) ||
                string.IsNullOrWhiteSpace(projectNameEditBox.Text) ||
                string.IsNullOrWhiteSpace(defaultAssigneeEditBox.Text))
            {
                MessageBox.Show("Please fill in all Azure DevOps configuration fields.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                _config = new AzureDevOpsConfig
                {
                    OrganizationUrl = organizationUrlEditBox.Text,
                    ProjectName = projectNameEditBox.Text,
                    DefaultAssignee = defaultAssigneeEditBox.Text
                };

                _credentialService.SaveAzureDevOpsConfig(_config);
                _azureDevOpsService = new AzureDevOpsService(_config);

                MessageBox.Show("Azure DevOps configuration saved successfully.", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving configuration: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            savePAT();
        }

        private void savePAT()
        {
            string enteredPat = patEditBox.Text;
            if (string.IsNullOrWhiteSpace(enteredPat))
            {
                MessageBox.Show("Please enter a valid PAT.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (enteredPat.Contains("●"))
            {
                MessageBox.Show("Please enter a new PAT, not the masked one.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                _credentialService.SavePat(enteredPat);
                patEditBox.Text = new string('●', 8);
                MessageBox.Show("PAT saved successfully.", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving PAT: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void createBug_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_azureDevOpsService == null)
                {
                    MessageBox.Show("Please configure Azure DevOps settings first.", "Configuration Required",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string pat = _credentialService.GetPat();
                if (string.IsNullOrEmpty(pat))
                {
                    MessageBox.Show("Please enter and save your PAT first.", "Missing PAT",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var mail = _emailService.GetSelectedEmail();
                if (mail == null)
                {
                    MessageBox.Show("Please select a mail item.", "No Selection",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string title = mail.Subject;
                string description = _emailService.CleanDescription(mail.Body);

                var result = await _azureDevOpsService.CreateBugAsync(title, description, pat);
                
                // Create the work item URL
                string workItemUrl = $"{_config.OrganizationUrl}/{_config.ProjectName}/_workitems/edit/{result.Id}";
                
                // Show message box with clickable link
                using (Form popup = new Form())
                {
                    popup.Text = "Bug Created Successfully";
                    popup.StartPosition = FormStartPosition.CenterScreen;
                    popup.Width = 400;
                    popup.Height = 150;

                    LinkLabel link = new LinkLabel
                    {
                        Text = $"Bug #{result.Id} created successfully. Click here to open.",
                        Width = 350,
                        Location = new System.Drawing.Point(25, 20),
                        AutoSize = true
                    };
                    
                    link.LinkClicked += (s, ev) => 
                    {
                        System.Diagnostics.Process.Start(workItemUrl);
                    };

                    Button closeButton = new Button
                    {
                        Text = "Close",
                        DialogResult = System.Windows.Forms.DialogResult.OK,
                        Location = new System.Drawing.Point(150, 60)
                    };

                    popup.Controls.AddRange(new Control[] { link, closeButton });
                    popup.AcceptButton = closeButton;
                    popup.ShowDialog();
                }
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
