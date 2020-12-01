using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [UsingHeaderNames(false)]
    [StartingFromRow(2)]
    public class OrdinalPropertiesColumnLetterAttributeTestClass
    {
        [MapsToColumnLetter("C")]
        public string ColumnC { get; set; }

        [MapsToColumnLetter("A")]
        public string ColumnA { get; set; }

        [MapsToColumnLetter("B")]
        public string ColumnB { get; set; }
    }
}