namespace Doxa.Labs.Excel.Models
{
    public class Cell
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public dynamic Value { get; set; }
        public dynamic ColumnWidth { get; set; }
    }
}
