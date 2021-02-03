using System;
using System.Collections.Generic;

namespace ExcelToEnumerable
{
    internal class PropertySetter
    {
        public Action<object, object> Setter { get; set; }

        public Type Type { get; set; }
        public string ColumnName { get; set; }

        public string PropertyName { get; set; }

        public IList<ExcelCellValidator> Validators { get; set; }

        /// <summary>
        ///     This will only be populated if we need to get value from the collection,
        ///     e.g. we need to check for uniqueness
        /// </summary>
        public Func<object, object> Getter { get; set; }

        public Func<object, object> PropertyMapping { get; internal set; }
        public bool RelaxedNumberMatching { get; set; }
    }
}