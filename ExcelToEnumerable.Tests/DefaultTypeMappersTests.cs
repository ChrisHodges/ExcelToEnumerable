using System;
using ExcelToEnumerable.Tests.TestClasses;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class DefaultTypeMappersTests
    {
        [Fact]
        public void EnumTypeMapperWorksWithNonAlphaNumericCharacters()
        {
            var enumTypeMapper =
                DefaultTypeMappers.CreateEnumTypeMapper(typeof(EnumsTestClass.ParseFromStringsEnum));
            enumTypeMapper("Forward /  SlashTest").Should().Be(EnumsTestClass.ParseFromStringsEnum.ForwardSlashTest);
        }
        
        [Fact]
        public void CreatesEnumTypeMapper()
        {
            var enumTypeMapper =
                DefaultTypeMappers.CreateEnumTypeMapper(typeof(EnumsTestClass.ParseFromStringsEnum));

            enumTypeMapper("Value1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value 1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value_1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value-1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value-2").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value2);
            enumTypeMapper(2).Should().Be(EnumsTestClass.ParseFromStringsEnum.Value2);
            enumTypeMapper("").Should().Be(default(EnumsTestClass.ParseFromStringsEnum));
            enumTypeMapper(null).Should().Be(default(EnumsTestClass.ParseFromStringsEnum));

            Action parseInvalidString = () => { enumTypeMapper("INVALID"); };
            parseInvalidString.Should().ThrowExactly<InvalidCastException>();
            
            Action parseInvalidInt = () => { enumTypeMapper(4); };
            parseInvalidInt.Should().ThrowExactly<InvalidCastException>();

            Action parseInvalidType = () => { enumTypeMapper(1.5M); };
            parseInvalidType.Should().ThrowExactly<InvalidCastException>();
        }
        
        [Fact]
        public void CreatesNullableEnumTypeMapper()
        {
            var enumTypeMapper =
                DefaultTypeMappers.CreateEnumTypeMapper(typeof(EnumsTestClass.ParseFromStringsEnum?));

            enumTypeMapper("Value1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value 1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value_1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value-1").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            enumTypeMapper("value-2").Should().Be(EnumsTestClass.ParseFromStringsEnum.Value2);
            enumTypeMapper(2).Should().Be(EnumsTestClass.ParseFromStringsEnum.Value2);
            enumTypeMapper("").Should().Be(null);
            enumTypeMapper(null).Should().Be(null);

            Action parseInvalidString = () => { enumTypeMapper("INVALID"); };
            parseInvalidString.Should().ThrowExactly<InvalidCastException>();
            
            Action parseInvalidInt = () => { enumTypeMapper(4); };
            parseInvalidInt.Should().ThrowExactly<InvalidCastException>();

            Action parseInvalidType = () => { enumTypeMapper(1.5M); };
            parseInvalidType.Should().ThrowExactly<InvalidCastException>();
        }
        
        [Fact]
        public void NullableEnumTypeMapperParsesInts()
        {
            var enumTypeMapper =
                DefaultTypeMappers.CreateEnumTypeMapper(typeof(EnumsTestClass.ParseFromIntsEnum?));

            enumTypeMapper(10).Should().Be(EnumsTestClass.ParseFromIntsEnum.Ten);
        }
    }
}