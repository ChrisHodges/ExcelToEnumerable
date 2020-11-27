using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [StartingFromRow(2)]
    [UsingSheet(1)]
    public class WorksheetByNumberAttributeTestClass
    {
        public string String { get; set; }
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}