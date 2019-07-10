using Xunit;

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
            Output output = new Output("sample");

            // Assert
            Assert.Equal(data, output.Log());
        }
    }
}
