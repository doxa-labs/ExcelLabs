using System;

namespace Excel.Labs
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
