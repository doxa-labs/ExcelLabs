using System;
using System.IO;
using System.Collections.Generic;
// Excel
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace Doxa.Labs.Excel.Models
{
    public class Output
    {
        // Excel
        private readonly ExcelApp.Application _app;
        private readonly ExcelApp.Workbook _workbook;
        // Excel - set from outside
        private readonly ExcelApp.Worksheet _worksheet;

        // Class Library
        public readonly string Title;
        public readonly string FilePath;
        public readonly string Extension;
        public Output(string title, string path, Extension extension)
        {
            // set title
            Title = title;

            // set extension
            switch (extension)
            {
                case Models.Extension.Xls:
                    Extension = ".xls";
                    break;
                case Models.Extension.Xlsx:
                    Extension = ".xlsx";
                    break;
            }

            // set path
            FilePath = Path.Combine(path, @"Files\" + title + Extension);

            // create excel app
            _app = new ExcelApp.Application();
            _workbook = _app.Workbooks.Add();
            _worksheet = (ExcelApp.Worksheet)_workbook.Worksheets[1];
        }

        public void SaveExcelFile(List<Cell> cells)
        {
            try
            {
                foreach (Cell item in cells)
                {
                    _worksheet.Cells[item.RowIndex, item.ColumnIndex] = item.Value;
                }

                // save and close
                _workbook.SaveAs(FilePath);
                _workbook.Close(true, Type.Missing, Type.Missing);
                _app.Quit();

                // clean
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {

            }
        }

        private static int CalculateStringWidth(string text)
        {
            return text.Length + 1;
        }
    }
}
