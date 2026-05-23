using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace OutlookAddIn.Services
{
    public interface IAzureDevOpsService
    {
        Task<WorkItem> CreateBugAsync(string title, string description, string pat);
        Task<WorkItem> CreateBugAsync(string title, string description, string pat, Dictionary<string, string> dynamicParameters);
        Task<WorkItem> CreateUserStoryAsync(string title, string description, string pat);
        Task<WorkItem> CreateUserStoryAsync(string title, string description, string pat, Dictionary<string, string> dynamicParameters);
    }
}