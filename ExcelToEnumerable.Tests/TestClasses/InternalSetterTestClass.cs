using System;

namespace ExcelToEnumerable.Tests
{
    public class InternalSetterTestClass
    {
        public string String { get; internal set; }
        public int? Int { get; internal set; }
        public DateTime? DateTime { get; internal set; }
        public double? Decimal { get; internal set; }
    }
}