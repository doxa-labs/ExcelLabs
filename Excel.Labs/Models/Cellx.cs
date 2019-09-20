namespace Doxa.Labs.Excel.Models
{
    /// <summary>
    /// OpenXml Cell Wrapper
    /// </summary>
    public class Cellx
    {
        /// <summary>
        /// Row Number
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// Column name like A, B, C, AA, AB
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// Your data as String or Number
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="value"></param>
        /// <param name="columnName"></param>
        public Cellx(int rowIndex, string value, string columnName = "-1")
        {
            RowIndex = rowIndex;
            Value = value;
            ColumnName = columnName;
        }
    }
}
