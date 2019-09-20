using System;

namespace Doxa.Labs.Excel.Models
{
    /// <summary>
    /// Throws when RowIndex is Zero
    /// </summary>
    public class ZeroIndexException : Exception
    {
        /// <summary>
        /// Throws when RowIndex is Zero
        /// </summary>
        /// <param name="message"></param>
        public ZeroIndexException(string message) : base(message)
        {
        }
    }

    /// <summary>
    /// Throws when Cell List is null
    /// </summary>
    public class NullCellListException : Exception
    {
        /// <summary>
        /// Throws when Cell List is null
        /// </summary>
        /// <param name="message"></param>
        public NullCellListException(string message) : base(message)
        {
        }
    }
}
