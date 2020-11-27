using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [StartingFromRow(2)]
    public class ShouldBeOneOfAttributeTestClass
    {
        [ShouldBeOneOf("abc123", "Steve", "Steven")]
        public string String { get; set; }
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}