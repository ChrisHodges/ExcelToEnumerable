using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [AggregateExceptions]
    [StartingFromRow(2)]
    public class RequiredAttributeTestClass
    {
        [Required]
        public string String { get; set; }
        
        [Required]
        public int? Int { get; set; }
        
        [Required]
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}