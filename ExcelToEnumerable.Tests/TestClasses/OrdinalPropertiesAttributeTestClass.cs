using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [UsingHeaderNames(false)]
    [StartingFromRow(2)]
    public class OrdinalPropertiesAttributeTestClass
    {
        [UsesColumnNumber(3)]
        public string ColumnC { get; set; }

        [UsesColumnNumber(1)]
        public string ColumnA { get; set; }

        [UsesColumnNumber(2)]
        public string ColumnB { get; set; }
    }
}