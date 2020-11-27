using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [UsingHeaderNames(false)]
    [StartingFromRow(2)]
    public class OrdinalPropertiesColumnLetterAttributeTestClass
    {
        [UsesColumnLetter("C")]
        public string ColumnC { get; set; }

        [UsesColumnLetter("A")]
        public string ColumnA { get; set; }

        [UsesColumnLetter("B")]
        public string ColumnB { get; set; }
    }
}