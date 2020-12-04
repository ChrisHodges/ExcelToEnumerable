using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [AggregateExceptions]
    [StartingFromRow(2)]
    public class NotNullAttributeTestClass
    {
        [NotNull]
        public string String { get; set; }
        
        [NotNull]
        public int? Int { get; set; }
        
        [NotNull]
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}