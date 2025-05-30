using System;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using OutlookAddIn1.Models;

namespace OutlookAddIn1.Services
{
    public class AzureDevOpsService : IAzureDevOpsService
    {
        private readonly AzureDevOpsConfig _config;

        public AzureDevOpsService(AzureDevOpsConfig config)
        {
            _config = config;
        }

        public async Task<WorkItem> CreateBugAsync(string title, string description, string pat)
        {
            var connection = new VssConnection(
                new Uri(_config.OrganizationUrl),
                new VssBasicCredential(string.Empty, pat));

            var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            var patchDocument = new JsonPatchDocument
            {
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Title", Value = title },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/Microsoft.VSTS.TCM.ReproSteps", Value = description },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.History", Value = "Created from Outlook email" },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.State", Value = "New" },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.AssignedTo", Value = _config.DefaultAssignee },
                new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Tags", Value = "Created-From-Outlook" }
            };

            return await witClient.CreateWorkItemAsync(patchDocument, _config.ProjectName, "Bug");
        }
    }
}