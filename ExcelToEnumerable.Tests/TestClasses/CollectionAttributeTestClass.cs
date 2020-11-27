using System;
using System.Collections.Generic;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    [UsingSheet("Sheet3CollectionExample")]
    public class CollectionAttributeTestClass
    {
        public string String { get; set; }
        public double? Int { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Decimal { get; set; }
        
        [MapFromColumns("CollectionColumn1","CollectionColumn2")]
        public List<string> Collection { get; set; }

        public List<int> IntCollection { get; set; }
    }
}