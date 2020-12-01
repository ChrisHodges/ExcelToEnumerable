using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [UsingHeaderNames(false)]
    [StartingFromRow(2)]
    public class OrdinalPropertiesAttributeTestClass
    {
        [MapsToColumnNumber(3)]
        public string ColumnC { get; set; }

        [MapsToColumnNumber(1)]
        public string ColumnA { get; set; }

        [MapsToColumnNumber(2)]
        public string ColumnB { get; set; }
    }

    [RelaxedNumberMatching(true)]
    public class RelaxedNumberMatchingAttributeTestClass
    {
        public int Int { get; set; }
    }
}