using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToEnumerable.Exceptions
{
    public class ExcelToEnumerableRowException : ExcelToEnumerableException
    {
        public ExcelToEnumerableRowException(object unmappedObject, string message, int rowNumber, IDictionary<string, object> rowValues, Exception innerException = null) : base(message, innerException)
        {
            RowNumber = rowNumber;
            RowValues = rowValues.ToDictionary(x => x.Key, x => x.Value.ToString());
            UnmappedObject = unmappedObject;
        }

        public int RowNumber { get; }
        public Dictionary<string, string> RowValues { get; }
        
        public object UnmappedObject { get; set; }
    }
}