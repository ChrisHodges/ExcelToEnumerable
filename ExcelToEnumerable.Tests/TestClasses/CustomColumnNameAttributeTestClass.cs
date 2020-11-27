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
        
        [UsesColumnNamed("IntCustomName")]
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}