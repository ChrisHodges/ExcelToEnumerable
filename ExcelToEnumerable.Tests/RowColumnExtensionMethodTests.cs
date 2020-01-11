using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class RowColumnExtensionMethodTests
    {
        [Fact]
        public void ToRowNameWorks1()
        {
            var result = 1.ToColumnName();
            result.Should().Be("A");

            result = 26.ToColumnName();
            result.Should().Be("Z");

            result = 27.ToColumnName();
            result.Should().Be("AA");

            result = 52.ToColumnName();
            result.Should().Be("AZ");
        }
    }
}