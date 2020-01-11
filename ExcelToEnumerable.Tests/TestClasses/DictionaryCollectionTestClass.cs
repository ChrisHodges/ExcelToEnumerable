using System;
using System.Collections.Generic;

namespace ExcelToEnumerable.Tests
{
    public class DictionaryCollectionTestClass
    {
        public string String { get; set; }
        public double? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
        public Dictionary<string, string> Collection { get; set; }

        public Dictionary<string, string> IntCollection { get; set; }
    }
}