using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Thrown when the mapper encounters an issue with an entire row of a spreadsheet. Exceptions relating to a specific cell
    /// throw an <see cref="ExcelToEnumerableCellException"/>, which inherits from this exception.
    /// </summary>
    public class ExcelToEnumerableRowException : ExcelToEnumerableException
    {
        internal ExcelToEnumerableRowException(object mappedObject, string message, int row,
            IDictionary<string, object> rowValues, Exception innerException = null) : base(message, innerException)
        {
            Row = row;
            RowValues = rowValues.ToDictionary(x => x.Key, x => x.Value?.ToString() ?? null);
            MappedObject = mappedObject;
        }

        /// <summary>
        /// The number of the spreadsheet row relating to the exception
        /// </summary>
        public int Row { get; }
        
        /// <summary>
        /// A column letter -> column value dictionary representing the values in spreadsheet row
        /// </summary>
        public Dictionary<string, string> RowValues { get; }
        
        /// <summary>
        /// The object that the mapper is attempting to map values to.
        /// </summary>
        public object MappedObject { get; set; }
    }
}