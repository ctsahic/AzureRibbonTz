using System.IO;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1.Services
{
    public interface IAzureDevOpsService
    {
        Task<WorkItem> CreateWorkItemAsync(string title, string description, string pat, string workItemType);
        Task AttachFilesToWorkItemAsync(int workItemId, Outlook.Attachments attachments, string pat);
        Task<WorkItem> UpdateWorkItemAsync(int workItemId, string comment, string pat); // Add this method
    }
}