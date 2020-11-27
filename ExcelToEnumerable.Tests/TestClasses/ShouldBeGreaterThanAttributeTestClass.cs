using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [StartingFromRow(2)]
    public class ShouldBeGreaterThanAttributeTestClass
    {
        public string String { get; set; }
        
        [ShouldBeGreaterThan(1)]
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}