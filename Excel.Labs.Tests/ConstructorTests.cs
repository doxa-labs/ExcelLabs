// Xunit - NuGet
using Xunit;
// Self
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.Tests
{
    public class ConstructorTests
    {
        [Theory]
        [InlineData("Excel.Labs", "path", Extension.Xls, @"path\Excel.Labs.xls", ".xls")]
        [InlineData("Excel.Labs", "path", Extension.Xlsx, @"path\Excel.Labs.xlsx", ".xlsx")]
        [InlineData("Excel.Labs", @"path\foldername", Extension.Xlsx, @"path\foldername\Excel.Labs.xlsx", ".xlsx")]
        public void Constructor_Should_Init_Properly(string title, string path, Extension extension, string expectedPath, string expectedExtension)
        {
            // Arrange

            // Act
            ExcelLabs excel = new ExcelLabs(title, path, extension);

            // Assert
            Assert.Equal(title, excel.Title);
            Assert.Equal(expectedPath, excel.FilePath);
            Assert.Equal(expectedExtension, excel.Extension);
        }

        [Fact]
        public void Constructor_Worksheet_Should_Not_Be_Null()
        {
            // Act
            ExcelLabs excel = new ExcelLabs("Excel.Labs", "path", Extension.Xls);

            // Assert
            Assert.NotNull(excel._worksheet);
        }

        [Fact]
        public void Constructor_Workbook_Should_Not_Be_Null()
        {
            // Act
            ExcelLabs excel = new ExcelLabs("Excel.Labs", "path", Extension.Xls);

            // Assert
            Assert.NotNull(excel._workbook);
        }
    }
}
