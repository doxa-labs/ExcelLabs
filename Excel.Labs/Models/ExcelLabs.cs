using System.IO;
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
            try
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
                    spreadsheetDocument.Close();
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
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