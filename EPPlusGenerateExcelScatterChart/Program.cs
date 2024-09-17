using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Drawing;

class Program
{
    static void Main(string[] args)
    {
        // Set the license context to non-commercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Define the output path for Windows environment
        string directoryPath = @"C:\app\output";  // Use a path appropriate for your Windows environment
        string filePath = Path.Combine(directoryPath, "FootballReportEPPlus.xlsx");

        // Ensure the directory exists
        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        using (ExcelPackage package = new ExcelPackage())
        {
            // Create a new worksheet in the Excel file
            var worksheet = package.Workbook.Worksheets.Add("Football Report");

            // Merge and format header rows
            worksheet.Cells["A1:E1"].Merge = true;
            worksheet.Cells["A1:E1"].Value = "Football Team Performance";
            worksheet.Cells["A1:E1"].Style.Font.Bold = true;
            worksheet.Cells["A1:E1"].Style.Font.Size = 16;
            worksheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A1:E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            // Add another header row below with individual column names
            worksheet.Cells["A2"].Value = "Team";
            worksheet.Cells["B2"].Value = "Wins";
            worksheet.Cells["C2"].Value = "Losses";
            worksheet.Cells["D2"].Value = "Draws";
            worksheet.Cells["E2"].Value = "Points";

            // Format header row
            using (var range = worksheet.Cells["A2:E2"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            // Add some data
            worksheet.Cells["A3"].Value = "Team A";
            worksheet.Cells["B3"].Value = 10;
            worksheet.Cells["C3"].Value = 5;
            worksheet.Cells["D3"].Value = 3;
            worksheet.Cells["E3"].Value = 33;

            worksheet.Cells["A4"].Value = "Team B";
            worksheet.Cells["B4"].Value = 12;
            worksheet.Cells["C4"].Value = 4;
            worksheet.Cells["D4"].Value = 2;
            worksheet.Cells["E4"].Value = 38;

            worksheet.Cells["A5"].Value = "Team C";
            worksheet.Cells["B5"].Value = 8;
            worksheet.Cells["C5"].Value = 6;
            worksheet.Cells["D5"].Value = 4;
            worksheet.Cells["E5"].Value = 28;

            // Auto-fit columns
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Apply borders to data cells
            using (var range = worksheet.Cells["A3:E5"])
            {
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            // Add a scatter chart
            var chart = worksheet.Drawings.AddChart("Football Performance", OfficeOpenXml.Drawing.Chart.eChartType.XYScatter);
            chart.Title.Text = "Team Performance";

            // Define X and Y series for the chart
            var series = chart.Series.Add(worksheet.Cells["B3:B5"], worksheet.Cells["E3:E5"]);
            series.Header = "Wins vs Points";

            // Position the chart
            chart.SetPosition(6, 0, 1, 0);
            chart.SetSize(600, 400);

            // Save the Excel file
            package.SaveAs(new FileInfo(filePath));

            Console.WriteLine("Excel file created successfully with colors, merged cells, and chart!");
        }
    }
}
