using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests
{
    [StartingFromRow(2)]
    public class AttributeTestClassWithRowNumber
    {
        [MapsToRowNumber]
        public int Row { get; set; }
        public string String { get; set; }
        public int? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
    }
}