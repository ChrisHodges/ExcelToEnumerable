using System.Collections.Generic;

namespace ExcelToEnumerable
{
    public class ExcelToEnumerableCollectionConfiguration
    {
        public string PropertyName { get; set; }
        public IEnumerable<string> ColumnNames { get; set; }
    }
}