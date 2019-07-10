using System;

namespace Excel.Labs.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            Output output = new Output("Hi");
            output.Log();
            Console.ReadLine();
        }
    }
}