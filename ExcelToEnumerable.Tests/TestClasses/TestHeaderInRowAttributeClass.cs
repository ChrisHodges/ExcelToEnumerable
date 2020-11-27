using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [HeaderOnRow(2)]
    [StartingFromRow(4)]
    [UsingSheet("HeaderOnRow2")]
    public class TestHeaderInRowAttributeClass
    {
        public string String { get; set; }
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}