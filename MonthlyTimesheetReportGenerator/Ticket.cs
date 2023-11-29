using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;

namespace AzureDevOpsWorkItems
{
    public class Ticket
    {
        public int? Id { get; set; }
        public string Type { get; set; }
        public string Status { get; set; }
        public string Iteration { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ClosedDate { get; set; }

        public override string ToString()
        {
            return $"{Type} {Id}- Status: {Status} Created: {CreatedDate}, Completed: {ClosedDate}";
        }

        public Ticket(WorkItem workitem)
        {
            var fields = workitem.Fields;

            this.Id = workitem.Id;
            this.Type = (string)workitem.Fields["System.WorkItemType"];
            this.Status = (string)workitem.Fields["System.State"];
            this.CreatedDate = (DateTime?)workitem.Fields["System.CreatedDate"];
            this.ClosedDate = (DateTime?)fields.GetValueOrDefault("Microsoft.VSTS.Common.ClosedDate");
            this.Iteration = (string)workitem.Fields["System.IterationPath"];
        }
    }
}
