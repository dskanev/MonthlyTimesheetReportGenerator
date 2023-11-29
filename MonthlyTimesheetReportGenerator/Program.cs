using MonthlyTimesheetReportGenerator;

namespace AzureDevOpsWorkItems
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var excelFile = ExcelService.GenerateEmptyExcelTimesheet())
            {
                ExcelService.PopulateVacation(
                new List<string>()
                    {
                        "10/11/2023",
                        "23/11/2023",
                        "24/11/2023"
                    }, 
                excelFile);

                var tickets = AzureDevopsService.GetTickets();

                ExcelService.PopulateTickets(tickets, excelFile);

                ExcelService.DownloadExcelFile(excelFile);
            }
        }
    }
}
