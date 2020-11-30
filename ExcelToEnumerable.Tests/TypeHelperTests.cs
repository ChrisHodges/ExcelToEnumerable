using System;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class TypeHelperTests
    {
        [Fact]
        public void NonNullableTypesWork()
        {
            typeof(int).IsNumeric().Should().BeTrue();
            
            typeof(double).IsNumeric().Should().BeTrue();
            
            typeof(decimal).IsNumeric().Should().BeTrue();
            
            typeof(float).IsNumeric().Should().BeTrue();
            
            typeof(long).IsNumeric().Should().BeTrue();
            
            typeof(string).IsNumeric().Should().BeFalse();
            
            typeof(DateTime).IsNumeric().Should().BeFalse();
        }
        
        [Fact]
        public void NullableTypesWork()
        {
            typeof(int?).IsNumeric().Should().BeTrue();
            
            typeof(double?).IsNumeric().Should().BeTrue();
            
            typeof(decimal?).IsNumeric().Should().BeTrue();
            
            typeof(float?).IsNumeric().Should().BeTrue();
            
            typeof(long?).IsNumeric().Should().BeTrue();

            typeof(DateTime?).IsNumeric().Should().BeFalse();
        }
    }
}