using System;
// Self
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            BaseOutput bo = new BaseOutput("title", "path", Extension.Xls);
            Console.WriteLine(bo.Extension);

            Console.ReadLine();
        }
    }
}