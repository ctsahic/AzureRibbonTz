using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using OutlookAddIn1.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1.Services
{
    public class AzureDevOpsService : IAzureDevOpsService
    {
        private readonly AzureDevOpsConfig _config;
        private readonly IEmailService _emailService;
        private VssConnection _connection;
        private string _currentPat;

        public AzureDevOpsService(AzureDevOpsConfig config, IEmailService emailService)
        {
            _config = config;
            _emailService = emailService;
        }

        private VssConnection GetConnection(string pat)
        {
            if (_connection == null || _currentPat != pat)
            {
                _connection = new VssConnection(
                    new Uri(_config.OrganizationUrl),
                    new VssBasicCredential(string.Empty, pat));
                _currentPat = pat;
            }
            return _connection;
        }

        public async Task AttachFilesToWorkItemAsync(int workItemId, Attachments attachments, string pat)
        {
            if (attachments == null || attachments.Count == 0)
                return;

            var witClient = GetConnection(pat).GetClient<WorkItemTrackingHttpClient>();
            string tempPath = Path.Combine(Path.GetTempPath(), "OutlookAttachments");
            
            try
            {
                // Create temp directory if it doesn't exist
                Directory.CreateDirectory(tempPath);

                foreach (Outlook.Attachment attachment in attachments)
                {
                    string tempFile = Path.Combine(tempPath, attachment.FileName);
                    
                    try
                    {
                        // Save attachment to temp file
                        attachment.SaveAsFile(tempFile);

                        // Upload attachment to Azure DevOps
                        using (FileStream fs = new FileStream(tempFile, FileMode.Open, FileAccess.Read))
                        {
                            // Specify the parameters explicitly to resolve ambiguity
                            var attachmentReference = await witClient.CreateAttachmentAsync(
                                uploadStream: fs,
                                fileName: attachment.FileName,
                                uploadType: "simple"
                            );

                            // Create patch operation to add attachment to work item
                            var patchDocument = new JsonPatchDocument
                            {
                                new JsonPatchOperation()
                                {
                                    Operation = Operation.Add,
                                    Path = "/relations/-",
                                    Value = new
                                    {
                                        rel = "AttachedFile",
                                        url = attachmentReference.Url,
                                        attributes = new { comment = "Attached from Outlook email" }
                                    }
                                }
                            };

                            // Update work item with attachment
                            await witClient.UpdateWorkItemAsync(patchDocument, workItemId, bypassRules: true);
                        }
                    }
                    finally
                    {
                        // Clean up temp file
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                }
            }
            finally
            {
                // Clean up temp directory
                try
                {
                    if (Directory.Exists(tempPath))
                    {
                        Directory.Delete(tempPath, true);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }


        public async Task<WorkItem> CreateWorkItemAsync(string title, string description, string pat, string workItemType)
        {
            var witClient = GetConnection(pat).GetClient<WorkItemTrackingHttpClient>();

            var patchDocument = new JsonPatchDocument
            {
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Title", Value = title },
                new JsonPatchOperation 
                { 
                    Operation = Operation.Add, 
                    Path = GetDescriptionField(workItemType), 
                    Value = description,
                    // Specify that we're sending HTML content
                    From = "text/html"
                },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.History", Value = "Created from Outlook email" },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Tags", Value = "Created-From-Outlook" }
            };

            // Only add assignee if provided
            if (!string.IsNullOrWhiteSpace(_config?.DefaultAssignee))
            {
                patchDocument.Add(
                    new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.AssignedTo", Value = _config.DefaultAssignee }
                );
            }

            // Use bypassRules = true to skip field validation - allows creation even with invalid/missing mandatory fields
            return await witClient.CreateWorkItemAsync(patchDocument, _config.ProjectName, workItemType, bypassRules: true);
        }

        public async Task<WorkItem> UpdateWorkItemAsync(int workItemId, string comment, string pat)
        {
            var witClient = GetConnection(pat).GetClient<WorkItemTrackingHttpClient>();

            var patchDocument = new JsonPatchDocument
            {
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/System.History",
                    Value = comment
                }
            };

            return await witClient.UpdateWorkItemAsync(patchDocument, workItemId);
        }

        private string GetDescriptionField(string workItemType)
        {
            return workItemType.Equals("Bug", StringComparison.OrdinalIgnoreCase) 
                ? "/fields/Microsoft.VSTS.TCM.ReproSteps" 
                : "/fields/System.Description";
        }

    }
}
