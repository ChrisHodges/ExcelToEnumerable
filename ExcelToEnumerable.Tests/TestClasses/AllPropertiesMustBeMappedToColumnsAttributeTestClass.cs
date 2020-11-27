using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [AllPropertiesMustBeMappedToColumns]
    public class AllPropertiesMustBeMappedToColumnsAttributeTestClass
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }
        
        public string NotOnSpreadsheet { get; set; } 
    }
}