using System;
// Excel
using ExcelApp = Microsoft.Office.Interop.Excel;
// Self
using Doxa.Labs.Excel.Models;
using System.Collections.Generic;

namespace Excel.Labs.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            string title = "Excel Labs NuGet";
            string sheetName = "Excel Sheet Name";
            string path = AppDomain.CurrentDomain.BaseDirectory;

            BaseOutput bo = new BaseOutput(title, path, Extension.Xls);
            Console.WriteLine(bo.Extension);
            Console.WriteLine(bo.FilePath);

            ExcelApp.Application app = new ExcelApp.Application();
            ExcelApp.Workbook workbook = app.Workbooks.Add(); // Missing.Value ?
            ExcelApp.Worksheet worksheet = workbook.Worksheets[1];
            worksheet.Name = $"{sheetName}";

            List<string> titleList = new List<string>() {
                "No"};

            for (int a = 1; a <= titleList.Count; a++)
            {
                worksheet.Cells[1, a].Value = titleList[a - 1];
            }

            worksheet.Cells[2, 9].Value = 5;
            worksheet.Cells[2, 10].Value = 5;

            try
            {
                workbook.SaveAs(path);
            }
            catch (Exception ex)
            {

            }
            workbook.Close(true, Type.Missing, Type.Missing);
            app.Quit();

            // New
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.ReadLine();
        }
    }
}