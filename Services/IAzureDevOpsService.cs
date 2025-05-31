using System.IO;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1.Services
{
    public interface IAzureDevOpsService
    {
        Task<WorkItem> CreateBugAsync(string title, string description, string pat);
        Task AttachFilesToWorkItemAsync(int workItemId, Outlook.Attachments attachments, string pat);
    }
}