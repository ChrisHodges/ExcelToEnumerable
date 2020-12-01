using System;
using System.Linq;
using ExcelToEnumerable.Tests.TestClasses;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class ExcelToEnumerableOptionsBuilderTests
    {
        private ExcelToEnumerableOptionsBuilder<TestClass> _subject;

        [Fact]
        public void CollectionConfigurationWorks()
        {
            var subject = new ExcelToEnumerableOptionsBuilder<CollectionTestClass>();
            subject.Property(x => x.Collection).MapFromColumns("ColumnA", "ColumnB");
            var options = subject.Build();
            var collectionConfiguration = options.CollectionConfigurations["Collection"];
            Console.WriteLine(collectionConfiguration.PropertyName); //Outputs "Collection";
            Console.WriteLine(collectionConfiguration.ColumnNames.First()); //Outputs "ColumnA";
            Console.WriteLine(collectionConfiguration.ColumnNames.Last()); //Outputs "ColumnN";
            collectionConfiguration.PropertyName.Should().Be("Collection");
            collectionConfiguration.ColumnNames.Should().BeEquivalentTo(new[] {"ColumnA", "ColumnB"});
        }

        [Fact]
        public void SetAggregateExceptionHandlingWorks()
        {
            _subject = new ExcelToEnumerableOptionsBuilder<TestClass>();
            _subject.AggregateExceptions();
            var result = _subject.Build();
            result.ExceptionHandlingBehaviour.Should().Be(ExceptionHandlingBehaviour.AggregateExceptions);
        }

        [Fact]
        public void RowNumbersAreOneBased()
        {
            _subject = new ExcelToEnumerableOptionsBuilder<TestClass>();
            _subject.StartingFromRow(2);
            var result = _subject.Build();
            result.StartRow.Should().Be(2);
        }

        [Fact]
        public void AddsAllPropertiesMustBeMappedToColumnsAttributeTestClassAttribute()
        {
            var builder = new ExcelToEnumerableOptionsBuilder<AllPropertiesMustBeMappedToColumnsAttributeTestClass>();
            var result = builder.Build();
            result.AllPropertiesOptionalByDefault.Should().BeFalse();
        }
        
        [Fact]
        public void Adds_AllColumnsMustBeMappedToProperties_Attribute()
        {
            var builder = new ExcelToEnumerableOptionsBuilder<AllColumnsMustBeMappedToPropertiesAttributeTestClass>();
            var result = builder.Build();
            result.IgnoreColumnsWithoutMatchingProperties.Should().BeFalse();
        }
        
        [Fact]
        public void Adds_MapsToColumnNumber_Attribute()
        {
            var builder = new ExcelToEnumerableOptionsBuilder<OrdinalPropertiesAttributeTestClass>();
            var result = builder.Build();
            result.CustomHeaderNumbers["ColumnA"].Should().Be(0);
            result.CustomHeaderNumbers["ColumnB"].Should().Be(1);
            result.CustomHeaderNumbers["ColumnC"].Should().Be(2);
        }
        
        [Fact]
        public void Adds_Unique_AttributeJustOnce()
        {
            var builder = new ExcelToEnumerableOptionsBuilder<UniqueAttributeTestClass>();
            var result = builder.Build();
            result.UniqueProperties.Count.Should().Be(2);
        }
        
        [Fact]
        public void Adds_RelaxedNumericMatching_Attribute()
        {
            var builder = new ExcelToEnumerableOptionsBuilder<RelaxedNumberMatchingAttributeTestClass>();
            var result = builder.Build();
            result.RelaxedNumberMatching.Should().BeTrue();
        }
    }
}