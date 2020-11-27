using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [AllColumnsMustBeMappedToProperties(true)]
    public class AllColumnsMustBeMappedToPropertiesAttributeTestClass
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }
    }
}