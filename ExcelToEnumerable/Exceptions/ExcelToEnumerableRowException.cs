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
        internal ExcelToEnumerableRowException(object unmappedObject, string message, int rowNumber,
            IDictionary<string, object> rowValues, Exception innerException = null) : base(message, innerException)
        {
            RowNumber = rowNumber;
            RowValues = rowValues.ToDictionary(x => x.Key, x => x.Value?.ToString() ?? null);
            UnmappedObject = unmappedObject;
        }

        /// <summary>
        /// The number of the spreadsheet row relating to the exception
        /// </summary>
        public int RowNumber { get; }
        
        /// <summary>
        /// A column letter -> column value dictionary representing the values in spreadsheet row
        /// </summary>
        public Dictionary<string, string> RowValues { get; }
        
        /// <summary>
        /// The object that the mapper is attempting to map values to.
        /// </summary>
        public object UnmappedObject { get; set; }
    }
}