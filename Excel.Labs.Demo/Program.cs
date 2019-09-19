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

            
            List<LabsCell> cells = new List<LabsCell>();
            /*
            for (int i = 1; i < 20; i++)
            {
                cells.Add(new LabsCell(i, i, i));
            }
            */

            // save excel file
            excel.Save(null);

            Console.ReadLine();
        }
    }
}