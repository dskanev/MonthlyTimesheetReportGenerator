using AzureDevOpsWorkItems;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.DataValidation;

namespace MonthlyTimesheetReportGenerator
{
    public static class ExcelService
    {
        public static void PopulateVacation(List<string> bookedVacation, ExcelPackage excelFile)
        {
            var worksheet = excelFile.Workbook.Worksheets["TimeSheet"];

            int startRow = 6;
            for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
            {
                if (bookedVacation.Contains(worksheet.Cells[row, 2].Text))
                {
                    worksheet.Cells[row, 3].Value = "VACATION";
                    worksheet.Cells[row, 4].Value = "VACATION";
                    worksheet.Cells[row, 5].Value = "VACATION";
                    worksheet.Cells[row, 6].Value = "VACATION";
                    worksheet.Cells[row, 7].Value = 0;
                }
            }
        }

        public static void PopulateTickets(List<Ticket> tickets, ExcelPackage excelFile)
        {
            foreach (var ticket in tickets.OrderBy(x => x.Id))
            {
                ExcelService.WriteTicketInNextEmptyRow(ticket, excelFile);
            }
        }

        private static void WriteTicketInNextEmptyRow(Ticket ticket, ExcelPackage excelFile)
        {
            var worksheet = excelFile.Workbook.Worksheets["TimeSheet"];
            string ticketStart = string.Empty;
            string ticketEnd = string.Empty;
            var isFilled = false;

            if (ticket.CreatedDate != null)
            {
                ticketStart = ticket.CreatedDate.Value.ToString("dd/MM/yyyy");
            }

            if (ticket.ClosedDate != null)
            {
                ticketEnd = ticket.ClosedDate.Value.ToString("dd/MM/yyyy");
            }

            // Assuming the dates start from row 6 and are in the second column (B)
            int startRow = 6; // Adjust if your starting row is different
            while (!isFilled)
            {
                for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 2].Text == ticketEnd)
                    {
                        var result = TryFitTicket(ticket, worksheet, row, "COMPLETED");
                        isFilled = result;
                    }
                    else if (worksheet.Cells[row, 2].Text == ticketStart)
                    {
                        var result = TryFitTicket(ticket, worksheet, row, "NEW");
                        isFilled = result;
                    }
                }
            }
        }

        private static bool TryFitTicket(Ticket ticket, ExcelWorksheet worksheet, int row, string status)
        {
            while (true)
            {
                if (worksheet.Cells[row, 3].Value == null)
                {
                    worksheet.Cells[row, 3].Value = ticket.Iteration;

                    if (string.Equals(ticket.Type, "Task"))
                    {
                        worksheet.Cells[row, 5].Value = $"{ticket.Id}";
                    }
                    else
                    {
                        worksheet.Cells[row, 4].Value = $"{ticket.Id}";
                    }

                    worksheet.Cells[row, 6].Value = status;
                    worksheet.Cells[row, 7].Value = 8;

                    return true;
                }
                else
                {
                    row++;
                }
            }
        }

        public static void PopulatePublicHolidays(List<string> bookedVacation, ExcelPackage excelFile)
        {
            var worksheet = excelFile.Workbook.Worksheets["TimeSheet"];

            int startRow = 6;
            for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
            {
                if (bookedVacation.Contains(worksheet.Cells[row, 2].Text))
                {
                    worksheet.Cells[row, 4].Value = "PUBLIC HOLIDAY";

                    worksheet.Cells[row, 7].Value = 8;
                }
            }
        }

        public static ExcelPackage GenerateEmptyExcelTimesheet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();

            var worksheet = AddWorksheet(package, "TimeSheet");

            CreateHeader(worksheet, "Denislav Kanev");

            CreateWorkingDaysAndGetLastRow(worksheet, out int lastRow);

            AddLastRowCalculation(worksheet, lastRow);

            StyleWorksheet(worksheet, lastRow);

            Console.WriteLine("Excel file generated successfully.");
            return package;
        }

        private static ExcelWorksheet AddWorksheet(ExcelPackage package, string name)
        {
            return package.Workbook.Worksheets.Add(name);
        }

        private static void CreateHeader(ExcelWorksheet worksheet, string employeeName)
        {
            // 1st Row
            worksheet.Cells["B3:G3"].Merge = true;
            worksheet.Cells["B3"].Value = $"KeyExpert: {employeeName}";

            // 2nd Row
            worksheet.Cells["B4:G4"].Merge = true;
            worksheet.Cells["B4"].Value = $"Month: {DateTime.Now.ToString("MMMM")}";

            // 3rd Row - Column Names
            string[] columnNames = { "Day", "Sprint Number", "PBI Reference", "Task Reference", "Status(New/Continued/Completed)", "Hours per day" };
            for (int i = 0; i < columnNames.Length; i++)
            {
                worksheet.Cells[5, i + 2].Value = columnNames[i]; // Starting from column B (index 2)
            }
        }

        private static void CreateWorkingDaysAndGetLastRow(ExcelWorksheet worksheet, out int lastRow)
        {
            int currentRow = 6;
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            for (DateTime date = firstDayOfMonth; date <= lastDayOfMonth; date = date.AddDays(1))
            {
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    worksheet.Cells[currentRow, 2].Value = date.ToString("dd/MM/yyyy"); // Column B
                    currentRow++;
                }
            }

            lastRow = currentRow;
        }

        private static void AddLastRowCalculation(ExcelWorksheet worksheet, int lastRow)
        {
            worksheet.Cells[lastRow, 2, lastRow, 6].Merge = true;
            worksheet.Cells[lastRow, 2].Value = "Total hours for the month";
            worksheet.Cells[lastRow, 7].Formula = $"SUM(G6:G{lastRow - 1})";
        }

        private static void StyleWorksheet(ExcelWorksheet worksheet, int lastRow)
        {
            // AutoFit Columns
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Apply thin borders to all cells
            var allCells = worksheet.Cells[worksheet.Dimension.Address];
            allCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            worksheet.Cells["B3:G3"].Style.Font.Bold = true;
            worksheet.Cells["B4:G4"].Style.Font.Bold = true;
            worksheet.Cells["B5:G5"].Style.Font.Bold = true;

            worksheet.Cells[lastRow, 2, lastRow, 7].Style.Font.Bold = true;
        }

        public static void DownloadExcelFile(ExcelPackage package)
        {
            // Save Excel File
            string currentDirectory = Directory.GetCurrentDirectory();
            string fileName = $"MITA Monthly report Denislav Kanev {DateTime.Now.ToString("MMMM")}.xlsx";
            string fullPath = Path.Combine(currentDirectory, fileName);

            FileInfo fileInfo = new FileInfo(fullPath);

            package.SaveAs(fileInfo);
        }
    }
}
