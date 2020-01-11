using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class ExcelToEnumerableMapperTests
    {
        private ExcelToEnumerableMapper<CollectionTestClass> _subject;

        [Fact]
        public void CollectionWorksCorrectly()
        {
            _subject = new ExcelToEnumerableMapper<CollectionTestClass>();
        }
    }
}