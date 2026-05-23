using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using OutlookAddIn.Models;

namespace OutlookAddIn.Services
{
    public interface IAzureDevOpsService
    {
        Task<WorkItem> CreateBugAsync(string title, string description, string pat, Dictionary<string, string> dynamicParameters = null, List<EmailAttachment> attachments = null);
        Task<WorkItem> CreateUserStoryAsync(string title, string description, string pat, Dictionary<string, string> dynamicParameters = null, List<EmailAttachment> attachments = null);
    }
}