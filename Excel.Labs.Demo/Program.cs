using System;
using System.Collections.Generic;
// Self
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            string title = "Excel Labs NuGet";
            string path = AppDomain.CurrentDomain.BaseDirectory;

            Output bo = new Output(title, path, Extension.Xls);
            // Console.WriteLine(bo.Extension);
            // Console.WriteLine(bo.FilePath);

            List<Cell> cells = new List<Cell>();

            for (int i = 1; i < 20; i++)
            {
                Cell c = new Cell() { RowIndex = 1, ColumnIndex = i, Value = i };
                cells.Add(c);
            }

            // save excel
            bo.SaveExcelFile(cells);

            Console.ReadLine();
        }
    }
}