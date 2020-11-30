using System;
using System.Globalization;
using System.Threading;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class RelaxedNumericConvertTests
    {
        [Fact]
        public void ConvertsToDouble()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            RelaxedNumericConvert.ToDouble("12345678").Should().Be(12345678);
            RelaxedNumericConvert.ToDouble("x123456789.222x").Should().Be(123456789.222);
            RelaxedNumericConvert.ToDouble("x -123,456,789.222 x").Should().Be(-123456789.222);
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            RelaxedNumericConvert.ToDouble("x -123.456.789,222 x").Should().Be(-123456789.222);

            Action action = () =>
            {
                RelaxedNumericConvert.ToDouble("No numeric content");
            };
            action.Should().ThrowExactly<FormatException>();
        }

        [Fact]
        public void SelectsFirstNumericPattern()
        {
            RelaxedNumericConvert.ToDouble("a string 12345678 then a break 987654 then another pattern").Should().Be(12345678);
        }
    }
}