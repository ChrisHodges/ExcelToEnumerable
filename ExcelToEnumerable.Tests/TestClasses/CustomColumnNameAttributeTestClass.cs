using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [StartingFromRow(2)]
    [EndingWithRow(2)]
    [UsingSheet("CustomColumnName")]
    public class CustomColumnNameAttributeTestClass
    {
        public string String { get; set; }
        
        [MapsToColumnNamed("IntCustomName")]
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }

    [StartingFromRow(2)]
    [EndingWithRow(2)]
    [UsingSheet("CustomColumnName")]
    public class CustomColumnNameTestClass
    {
        public string String { get; set; }
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}