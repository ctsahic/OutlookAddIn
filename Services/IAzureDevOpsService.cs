using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace OutlookAddIn.Services
{
    public interface IAzureDevOpsService
    {
        Task<WorkItem> CreateBugAsync(string title, string description, string pat);
    }
}