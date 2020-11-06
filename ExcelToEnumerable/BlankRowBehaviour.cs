namespace ExcelToEnumerable
{
    /// <summary>
    /// Defines what ExcelToEnumerable should do when it encounters a blank row in the spreadsheet
    /// e.g.
    ///  testSpreadsheetLocation.ExcelToEnumerable	&lt;SomeClass&gt; (x => .BlankRowBehaviour(BlankRowBehaviour.ThrowException))
    /// </summary>
    public enum BlankRowBehaviour
    {
        /// <summary>
        /// Ignores the blank row, but continues reading the spreadsheet
        /// </summary>
        Ignore,
        
        /// <summary>
        /// Throws an ExcelToEnumerableRowException
        /// </summary>
        ThrowException,
        
        /// <summary>
        /// Stops reading the spreadsheet, ignoring any rows after the blank row
        /// </summary>
        StopReading,
        
        /// <summary>
        /// Creates an instance of type T, where T is the Enumerable type being mapped
        /// </summary>
        CreateEntity
    }
}