using AzureDevOpsWorkItems;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyTimesheetReportGenerator
{
    public static class AzureDevopsService
    {
        public static List<Ticket> GetTickets()
        {
            var tickets = new List<Ticket>();

            try
            {
                string orgUrl = "https://dev.azure.com/XXX";
                string personalAccessToken = "XXX";

                VssConnection connection = new VssConnection(new Uri(orgUrl), new VssBasicCredential(string.Empty, personalAccessToken));

                var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

                int year = DateTime.Now.Year;
                int month = DateTime.Now.Month;

                DateTime firstDayOfMonth = new DateTime(year, month, 1);
                DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

                string wiqlQuery = $@"
                Select [System.Id], [System.Title], [System.State], [System.WorkItemType]
                From WorkItems 
                Where [System.ChangedDate] >= '{firstDayOfMonth:yyyy-MM-dd}' 
                  And [System.ChangedDate] <= '{lastDayOfMonth:yyyy-MM-dd}' 
                  And [System.ChangedBy] = @Me
                Order By [System.ChangedDate] Desc";

                Wiql wiql = new Wiql()
                {
                    Query = wiqlQuery
                };

                WorkItemQueryResult workItemQueryResult = witClient.QueryByWiqlAsync(wiql).Result;

                if (workItemQueryResult.WorkItems.Any())
                {
                    List<WorkItem> workItems = witClient.GetWorkItemsAsync(workItemQueryResult.WorkItems.Select(wi => wi.Id)).Result;


                    foreach (var workItem in workItems.OrderBy(x => x.Id))
                    {
                        var ticket = new Ticket(workItem);
                        tickets.Add(ticket);
                        Console.WriteLine(ticket.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return tickets;
        }
    }
}
