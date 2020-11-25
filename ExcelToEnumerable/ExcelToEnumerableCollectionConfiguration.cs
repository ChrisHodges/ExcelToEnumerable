using System.Collections.Generic;

namespace ExcelToEnumerable
{
    internal class ExcelToEnumerableCollectionConfiguration
    {
        public string PropertyName { get; set; }
        
        public IEnumerable<string> ColumnNames { get; set; }
    }
}