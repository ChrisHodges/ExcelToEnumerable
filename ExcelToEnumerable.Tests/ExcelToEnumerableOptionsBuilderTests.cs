using System;
using System.Linq;
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
    }
}