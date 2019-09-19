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
            string path = AppDomain.CurrentDomain.BaseDirectory + @"Files\";

            ExcelLabs excel = new ExcelLabs(title, path, Extension.Xls);
            // Console.WriteLine(bo.Extension);
            Console.WriteLine(excel.FilePath);

            // create a cell list
            List<LabsCell> cells = new List<LabsCell>();

            // define row and column indexes then add your data
            cells.Add(new LabsCell(10, 20, "Your Value"));

            // add some data to cell list
            for (int i = 1; i < 20; i++)
            {
                cells.Add(new LabsCell(i, i, i));
            }

            // call save function with the cell list
            excel.Save(cells);

            Console.ReadLine();
        }
    }
}