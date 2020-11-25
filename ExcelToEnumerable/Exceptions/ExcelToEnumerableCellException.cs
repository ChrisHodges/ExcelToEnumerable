using System;
using System.Collections.Generic;

namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Thrown when the mapper fails to map a cell value to a property.
    /// </summary>
    public class ExcelToEnumerableCellException : ExcelToEnumerableRowException
    {
        internal ExcelToEnumerableCellException(object unmappedObject, int row, int column, object cellValue,
            string propertyName, IDictionary<string, object> rowValues,
            string validationMessage = null, Exception innerException = null,
            ExcelToEnumerableValidationCode? excelToEnumerableValidationCode = null)
            :
            base(unmappedObject,
                $"Unable to set value '{cellValue}' to property '{propertyName}' on row {row} column {(column + 1).ToColumnName()}. {validationMessage}",
                row,
                rowValues,
                innerException)
        {
            Column = (column + 1).ToColumnName();
            PropertyName = propertyName;
            CellValue = cellValue;
            ExcelToEnumerableValidationCode = excelToEnumerableValidationCode;
        }

        /// <summary>
        /// The column letter of the cell that caused the exception.
        /// </summary>
        public string Column { get; }
        
        /// <summary>
        /// The name of the property the mapper attempted to map to.
        /// </summary>
        public string PropertyName { get; }
        
        /// <summary>
        /// The value of the cell that was being mapped
        /// </summary>
        public object CellValue { get; }
        
        /// <summary>
        /// An optional <see cref="ExcelToEnumerable.Exceptions.ExcelToEnumerableValidationCode"/> enum identifying the validation rule that failed.
        /// </summary>
        public ExcelToEnumerableValidationCode? ExcelToEnumerableValidationCode { get; }
    }
}