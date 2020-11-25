using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable
{
    /// <summary>
    ///     Enum. Defines what ExcelToEnumerable should do when it encounters a blank row in the spreadsheet. Defaults to <c>ExcelToEnumerable.ThrowException</c>
    /// </summary>
    /// <example>
    /// <code>
    /// IEnumerable&lt;MyClass&gt; results = spreadsheetStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
    ///   .BlankRowBehaviour(BlankRowBehaviour.ThrowException)
    /// );
    /// </code>
    /// </example>
    public enum BlankRowBehaviour
    {
        /// <summary>
        ///     Ignores the blank row but continues reading the spreadsheet
        /// </summary>
        Ignore,

        /// <summary>
        ///     Stops reading the spreadsheet and throws an <see cref="ExcelToEnumerableRowException"/>
        /// </summary>
        ThrowException,

        /// <summary>
        ///     Stops reading the spreadsheet, ignoring any rows after the blank row
        /// </summary>
        StopReading,

        /// <summary>
        ///     Creates an instance of the mapped type with the default constructor and continues reading the spreadsheet
        /// </summary>
        CreateEntity
    }
}