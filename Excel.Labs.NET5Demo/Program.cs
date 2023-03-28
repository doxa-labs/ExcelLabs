using System;
using System.Collections.Generic;
// team
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.NET5Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            string title = "Excel Labs NuGet";
            string sheetName = "Simple and Fast";

            // TODO: create a folder named Files to run this demo
            // TODO: or, you may change the path
            // TODO: full path: C:\Users\...\ExcelLabs\Excel.Labs.NET5Demo\bin\Debug\net6.0\Files
            string path = AppDomain.CurrentDomain.BaseDirectory + @"Files\";

            // 1. create a cell list
            List<Cellx> cells = new();

            // 2. values as an array
            List<string> languages = new()
            {
                "Java", // A
                "C#", // B
                "Javascript", // C
                "Swift", // D
                "Php", // E
                "Python", // F
                "Go", // G
                "Swift", // H
                "", // I
                "", // J
                "", // K
                "Objective-C", // L
                "C++", // M
                "F#" // N
            };

            foreach (string lang in languages)
            {
                // no column name for ordered columns
                cells.Add(new Cellx(1, lang));
            }

            // 3. single value with column name
            cells.Add(new Cellx(2, "Fortran", "A"));
            cells.Add(new Cellx(2, "Cobol", "D"));
            cells.Add(new Cellx(2, "Pascal", "I"));

            // 4. single value without column name
            cells.Add(new Cellx(3, "Visual Studio"));
            cells.Add(new Cellx(3, "Webstorm"));
            cells.Add(new Cellx(3, "Xcode"));
            cells.Add(new Cellx(3, "Notepad"));

            // call save function
            ExcelLabs.SaveFile(title, path, sheetName, cells);

            Console.WriteLine("Done. Check the path now to see the Excel file.");
            Console.ReadLine();
        }
    }
}
