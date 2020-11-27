using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [StartingFromRow(2)]
    [AggregateExceptions]
    [UsingSheet("Sheet4DuplicateStringValue")]
    public class UniqueAttributeTestClass
    {
        [Unique]
        public string String { get; set; }
        
        [Unique]
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}