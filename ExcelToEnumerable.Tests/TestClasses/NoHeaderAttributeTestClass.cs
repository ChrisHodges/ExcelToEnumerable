using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests
{
    [UsingHeaderNames(false)]
    public class NoHeaderAttributeTestClass
    {
        public string ColumnA { get; set; }
        public int ColumnB { get; set; }
    }
}