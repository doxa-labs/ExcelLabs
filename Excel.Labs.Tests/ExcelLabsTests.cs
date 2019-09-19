using System.Collections.Generic;
// Xunit - NuGet
using Xunit;
// Self
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.Tests
{
    public class ExcelLabsTests
    {
        [Theory]
        [InlineData(0, 1, "Test Value")]
        [InlineData(1, 0, "Test Value")]
        [InlineData(0, 0, "Test Value")]
        public void Save_Throws_Exception_When_Index_Zero(int rowIndex, int columnIndex, dynamic value)
        {
            // Arrange
            ExcelLabs excel = new ExcelLabs("Excel.Labs", "path", Extension.Xls);

            // Act
            List<LabsCell> cells = new List<LabsCell>();
            cells.Add(new LabsCell(rowIndex, columnIndex, value));

            // Assert
            Assert.Throws<ZeroIndexException>(() => excel.Save(cells));
        }

        [Fact]
        public void Save_Throws_Exception_When_CellList_Null()
        {
            // Arrange
            ExcelLabs excel = new ExcelLabs("Excel.Labs", "path", Extension.Xls);

            // Assert
            Assert.Throws<NullCellListException>(() => excel.Save(null));
        }
    }
}