using System;

namespace Doxa.Labs.Excel.Models
{
    public class ZeroIndexException : Exception
    {
        public ZeroIndexException(string message) : base(message)
        {
        }
    }

    public class NullCellListException : Exception
    {
        public NullCellListException(string message) : base(message)
        {
        }
    }
}
