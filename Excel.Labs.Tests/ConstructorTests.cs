using Xunit;
// Self
using Doxa.Labs.Excel.Models;

namespace Excel.Labs.Tests
{
    public class ConstructorTests
    {
        [Fact]
        public void Constructor_Output_Should_Init_Text()
        {
            // Arrange
            string data = "sample";

            // Act
            ExcelLabs output = new ExcelLabs("sample", "path", Extension.Xls);

            // Assert
            Assert.Equal(data, output.Title);
        }

        [Fact]
        public void Constructor_Worksheet_Should_Not_Be_Null()
        {
            // Act
            ExcelLabs output = new ExcelLabs("sample", "path", Extension.Xls);

            // Assert
            Assert.NotNull(output._worksheet);
        }
    }
}
