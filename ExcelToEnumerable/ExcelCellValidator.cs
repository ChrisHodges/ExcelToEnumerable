using System;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable
{
    /// <summary>
    ///     A validator function, a message to display if the worksheet cell fails validation, and optionally a validation code
    ///     to uniquely identify the validation rule.
    /// </summary>
    public class ExcelCellValidator
    {
        /// <summary>
        ///     A human-readable message to display to a user, describing why a cell has failed validation.
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        ///     A validation function - accepts an object as it's single argument and returns a Boolean for pass/fail.
        /// </summary>
        public Func<object, bool> Validator { get; set; }

        /// <summary>
        ///     An option validation code to uniquely identify the validation rule.
        /// </summary>
        public ExcelToEnumerableValidationCode? ExcelToEnumerableValidationCode { get; set; }
    }
}