using System;
using System.Collections.Generic;

namespace ExcelToEnumerable.Tests
{
    public class CollectionTestClass
    {
        public string String { get; set; }
        public double? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
        public List<string> Collection { get; set; }

        public List<int> IntCollection { get; set; }
    }
}