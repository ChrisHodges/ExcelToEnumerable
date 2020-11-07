using System;
using System.Collections.Generic;

namespace ExcelToEnumerable.Exceptions
{
    public class ExcelToEnumerableCellException : ExcelToEnumerableRowException
    {
        /// <summary>
        /// </summary>
        /// <param name="row">
        ///     This is zero indexed!
        /// </param>
        /// <param name="column">
        ///     This is zero indexed!
        /// </param>
        /// <param name="cellValue"></param>
        /// <param name="propertyName"></param>
        /// <param name="validationMessage"></param>
        /// <param name="innerException"></param>
        public ExcelToEnumerableCellException(object unmappedObject, int row, int column, object cellValue,
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

        public string Column { get; }
        public string PropertyName { get; }
        public object CellValue { get; }
        public ExcelToEnumerableValidationCode? ExcelToEnumerableValidationCode { get; }
    }
}