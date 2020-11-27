using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [UsingSheet("2Columns")]
    [AllPropertiesMustBeMappedToColumns(true)]
    public class OptionalParametersAttributeTestClass
    {
        public string Name { get; set; }
        public decimal Fee1 { get; set; }
        
        [Optional]
        public decimal Fee2 { get; set; }
        public decimal Fee3 { get; set; }
    }
}