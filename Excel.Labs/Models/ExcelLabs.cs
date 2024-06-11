using System.IO;
using System.Linq;
using System.Collections.Generic;
// OpenXml - NuGet
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Doxa.Labs.Excel.Models
{
    /// <summary>
    /// Create Excel File Simple and Fast.
    /// </summary>
    public class ExcelLabs
    {
        /// <summary>
        /// Gets data as a Cellx List format and Save the Excel file as .xlsx
        /// </summary>
        /// <param name="title"></param>
        /// <param name="path"></param>
        /// <param name="sheetName"></param>
        /// <param name="cells"></param>
        public static void SaveFile(string title, string path, string sheetName, List<Cellx> cells)
        {
            // check for null cell list
            if (cells == null)
            {
                throw new NullCellListException("Cell List cannot be null.");
            }

            // check for rowindex == 0
            if (cells.Exists(a => a.RowIndex == 0))
            {
                throw new ZeroIndexException("RowIndex should be greater than 0. It starts from 1.");
            }

            // generate the full path
            string fullPath = Path.Combine(path, title + ".xlsx");

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook))
            {
                // add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };

                int tempRow = 0;
                Row row = new Row();
                foreach (Cellx item in cells)
                {
                    if (tempRow != item.RowIndex)
                    {
                        if (tempRow != 0)
                        {
                            sheetData.Append(row);
                        }

                        // init row with row index
                        row = new Row() { RowIndex = (uint)item.RowIndex };

                        // save row index
                        tempRow = item.RowIndex;
                    }

                    string cellReference = item.ColumnName + item.RowIndex;
                    // columnName = -1 for ordered columns
                    if (item.ColumnName == "-1")
                    {
                        cellReference = item.ColumnName;
                    }

                    row.Append(new Cell() { CellReference = cellReference, CellValue = new CellValue(item.Value), DataType = ResolveCellDataTypeOnValue(item.Value).Value });
                }

                // append last row
                sheetData.Append(row);

                sheets.Append(sheet);
                workbookpart.Workbook.Save();

                // close the document.
                // Close() is obsolete on DocumentFormat.OpenXml 3.0.3 due to some crash
                // Check for the details https://github.com/dotnet/Open-XML-SDK/releases/tag/v3.0.3
                //spreadsheetDocument.Close();

                // started using Dispose instead of Close with 3.0.3
                spreadsheetDocument.Dispose();
            }
        }

        /// <summary>
        /// Gets data as a Cellx List format and Save the Excel file as .xlsx + cleans potential not-allowed XML characters
        /// </summary>
        /// <param name="title"></param>
        /// <param name="path"></param>
        /// <param name="sheetName"></param>
        /// <param name="cells"></param>
        public static void SaveFileWithCleanXmlText(string title, string path, string sheetName, List<Cellx> cells)
        {
            // check for null cell list
            if (cells == null)
            {
                throw new NullCellListException("Cell List cannot be null.");
            }

            // check for rowindex == 0
            if (cells.Exists(a => a.RowIndex == 0))
            {
                throw new ZeroIndexException("RowIndex should be greater than 0. It starts from 1.");
            }

            // generate the full path
            string fullPath = Path.Combine(path, title + ".xlsx");

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook))
            {
                // add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };

                int tempRow = 0;
                Row row = new Row();
                foreach (Cellx item in cells)
                {
                    if (tempRow != item.RowIndex)
                    {
                        if (tempRow != 0)
                        {
                            sheetData.Append(row);
                        }

                        // init row with row index
                        row = new Row() { RowIndex = (uint)item.RowIndex };

                        // save row index
                        tempRow = item.RowIndex;
                    }

                    string cellReference = item.ColumnName + item.RowIndex;
                    // columnName = -1 for ordered columns
                    if (item.ColumnName == "-1")
                    {
                        cellReference = item.ColumnName;
                    }

                    row.Append(new Cell() { CellReference = cellReference, CellValue = new CellValue(CleanTextForXml(item.Value)), DataType = ResolveCellDataTypeOnValue(item.Value).Value });
                }

                // append last row
                sheetData.Append(row);

                sheets.Append(sheet);
                workbookpart.Workbook.Save();

                // close the document.
                // Close() is obsolete on DocumentFormat.OpenXml 3.0.3 due to some crash
                // Check for the details https://github.com/dotnet/Open-XML-SDK/releases/tag/v3.0.3
                //spreadsheetDocument.Close();

                // started using Dispose instead of Close with 3.0.3
                spreadsheetDocument.Dispose();
            }
        }

        /// <summary>
        /// Cleans potential not-allowed XML characters
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string CleanTextForXml(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "-";
            }

            try
            {
                string clean = System.Net.WebUtility.HtmlDecode(text);
                clean = new string(clean.Where(ch => System.Xml.XmlConvert.IsXmlChar(ch)).ToArray());

                return clean;
            }
            catch
            {
                return "-";
            }
        }

        /// <summary>
        /// Convert integers to string like 1 to A, 2 to B, 3 to C. You can write your text in different columns by using this function
        /// Find details https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static string ColumnIndexToColumnLetter(int columnIndex)
        {
            var index = columnIndex;
            var columnLetter = string.Empty;
            int mod;

            while (index > 0)
            {
                mod = (index - 1) % 26;
                columnLetter = (char)(65 + mod) + columnLetter;
                index = (index - mod) / 26;
            }

            return columnLetter;
        }

        /// <summary>
        /// Detects the correct value type
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Returns the correct value type</returns>
        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }
}