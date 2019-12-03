using System.Collections.Generic;
// Xunit - NuGet
using Xunit;
// Self
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.Tests
{
    public class ExcelLabsTests
    {
        /*
        [Theory]
        [InlineData(0, "Test Value", "A")]
        public void Save_Throws_Exception_When_RowIndex_Zero(int rowIndex, string value, string columnName)
        {
            // Arrange
            string title = "Title";
            string path = "Path";
            string sheetName = "Sheet Name";

            // Act
            List<Cellx> cells = new List<Cellx>();
            cells.Add(new Cellx(rowIndex, value, columnName));

            // Assert
            Assert.Throws<ZeroIndexException>(() => ExcelLabs.SaveFile(title, path, sheetName, cells));
        }
        */

        /*
        [Fact]
        public void Save_Throws_Exception_When_CellList_Null()
        {
            // Arrange
            string title = "Title";
            string path = "Path";
            string sheetName = "Sheet Name";

            // Act

            // Assert
            Assert.Throws<NullCellListException>(() => ExcelLabs.SaveFile(title, path, sheetName, null));
        }
        */
    }
}