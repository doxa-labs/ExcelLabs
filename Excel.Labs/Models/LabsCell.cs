namespace Doxa.Labs.Excel.Models
{
    public class LabsCell
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public dynamic Value { get; set; }

        public LabsCell(int rowIndex, int columnIndex, dynamic value)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            Value = value;
        }
    }
}
