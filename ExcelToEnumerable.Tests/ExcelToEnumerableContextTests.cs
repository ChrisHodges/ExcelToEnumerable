using System.Linq;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class ExcelToEnumerableContextTests
    {
        [Fact]
        public void FromRowConstructorHasCorrectNumberOfSetters()
        {
            var options = new ExcelToEnumerableOptions<TestClass2>();
            ExcelToEnumerableContext.Instance.CreateMapper(options);
            var fromRowConstructor = ExcelToEnumerableContext.Instance.GetFromRowConstructor(options);
            fromRowConstructor.Setters.Count().Should().Be(14);
        }

        [Fact]
        public void FromRowConstructorWithCollectionHasCorrectNumberOfSetters()
        {
            var options = new ExcelToEnumerableOptions<CollectionTestClass>();
        }
    }
}