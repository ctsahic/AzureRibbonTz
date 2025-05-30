using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace OutlookAddIn1.Services
{
    public interface IAzureDevOpsService
    {
        Task<WorkItem> CreateBugAsync(string title, string description, string pat);
    }
}