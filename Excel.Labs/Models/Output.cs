using System;

namespace Doxa.Labs.Excel.Models
{
    public class Output
    {
        private readonly string _text;
        public Output(string text)
        {
            _text = text;
        }

        public string Log()
        {
            Console.WriteLine(_text);
            return _text;
        }
    }
}
