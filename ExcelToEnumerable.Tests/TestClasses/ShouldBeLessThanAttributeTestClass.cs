using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [StartingFromRow(3)]
    public class ShouldBeLessThanAttributeTestClass
    {
        public string String { get; set; }
        
        [ShouldBeLessThan(2)]
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}