using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class ExcelToEnumerableOptionsBuilderTests
    {
        [Fact]
        public void SetAggregateExceptionHandlingWorks()
        {
            ExcelToEnumerableOptionsBuilder<TestClass> _subject;
            _subject = new ExcelToEnumerableOptionsBuilder<TestClass>();
            _subject.AggregateExceptions();
            var result = _subject.Build();
            result.ExceptionHandlingBehaviour.Should().Be(ExceptionHandlingBehaviour.AggregateExceptions);
        }

        [Fact]
        public void RowNumbersAreOneBased()
        {
            ExcelToEnumerableOptionsBuilder<TestClass> _subject;
            _subject = new ExcelToEnumerableOptionsBuilder<TestClass>();
            _subject.StartingFromRow(2);
            var result = _subject.Build();
            result.StartRow.Should().Be(2);
        }
        
        [Fact]
        public void DefaultRowNumberIsOne()
        {
            ExcelToEnumerableOptionsBuilder<TestClass> _subject;
            _subject = new ExcelToEnumerableOptionsBuilder<TestClass>();
            var result = _subject.Build();
            result.StartRow.Should().Be(1);
        }
    }
}