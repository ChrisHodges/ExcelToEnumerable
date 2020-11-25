namespace ExcelToEnumerable
{
    /// <summary>
    /// Enum. Determines how the mapper should respond to validation exceptions when reading the spreadsheet. Defaults to <c>ThrowOnFirstException</c>
    /// Not set directly, but via the fluent configuration-builder argument in one of the ExcelToEnumerable <see cref="ExtensionMethods"/>
    /// </summary>
    public enum ExceptionHandlingBehaviour
    {
        /// <summary>
        /// Will throw an exception and stop reading the spreadsheet as soon as the first exception is encountered.
        /// </summary>
        ThrowOnFirstException,
        /// <summary>
        /// Will read the entire spreadsheet then throw an AggregateException if any exceptions are encountered.
        /// </summary>
        AggregateExceptions,
        /// <summary>
        /// Will add exceptions to the IExcelToEnumerableOption ExceptionList property.
        /// </summary>
        LogExceptions
    }
}