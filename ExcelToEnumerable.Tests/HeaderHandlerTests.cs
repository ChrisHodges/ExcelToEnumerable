using System;
using ExcelToEnumerable.Exceptions;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class HeaderHandlerTests
    {
        [Fact]
        public void ThrowsExceptionForColumnNotMappedToProperty()
        {
            Action action = () =>
            {
                var normalisedPropertyNames = new[] {"a", "b", "c"};
                var normalisedHeaderNames = new[] {"a", "b", "c", "d"};
                var unmappedProperties = new string[0];
                var optionalProperties = new string[0];
                var ignoreColumnsWithoutMatchingProperties = false;
                HeaderHandler.ValidateColumnNames(normalisedPropertyNames, normalisedHeaderNames, unmappedProperties,
                    optionalProperties, ignoreColumnsWithoutMatchingProperties);
            };
            action.Should().ThrowExactly<ExcelToEnumerableInvalidHeaderException>().WithMessage("Missing properties: 'd'.");
        }
    }
}