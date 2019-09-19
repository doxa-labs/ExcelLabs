using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
// Excel - NuGet
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace Doxa.Labs.Excel.Models
{
    public class ExcelLabs
    {
        // Excel
        private readonly ExcelApp.Application _app;
        public readonly ExcelApp.Workbook _workbook;
        public readonly ExcelApp.Worksheet _worksheet;

        // Class Library
        public readonly string Title;
        public readonly string FilePath;
        public readonly string Extension;
        public ExcelLabs(string title, string path, Extension extension)
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
            FilePath = Path.Combine(path, title + Extension);

            // create excel app and worksheet
            _app = new ExcelApp.Application();
            _workbook = _app.Workbooks.Add();
            _worksheet = (ExcelApp.Worksheet)_workbook.Worksheets[1];
        }

        /// <summary>
        /// Gets the data as List<Cell> and Save as an Excel File
        /// </summary>
        /// <param name="cells"></param>
        public void Save(List<LabsCell> cells)
        {
            try
            {
                // check cellList
                if (cells == null)
                {
                    throw new NullCellListException("Cell List cannot be Null.");
                }

                foreach (LabsCell item in cells)
                {
                    // check for 0 index
                    if (item.RowIndex == 0 || item.ColumnIndex == 0)
                    {
                        throw new ZeroIndexException("RowIndex or ColumnIndex cannot be Zero.");
                    }

                    _worksheet.Cells[item.RowIndex, item.ColumnIndex] = item.Value;
                }

                // save
                _workbook.SaveAs(FilePath);
                
                // cleanup
                Marshal.FinalReleaseComObject(_worksheet);

                _workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(_workbook);

                _app.Quit();
                Marshal.FinalReleaseComObject(_app);

                // first run
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // second run
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
