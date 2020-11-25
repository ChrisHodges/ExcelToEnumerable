namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Thrown went data relating to an entire sheet is invalid, for example if duplicate values appear for a unique property.
    /// </summary>
    public class ExcelToEnumerableSheetException : ExcelToEnumerableException
    {
        internal ExcelToEnumerableSheetException(string message) : base(message)
        {
        }
    }
}