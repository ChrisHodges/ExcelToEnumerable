using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Tells the mapper to stop reading data at the specified 1-based row number. If negative, the mapper will read until the
    /// specified number of rows from the bottom of the spreadsheet.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class EndingWithRowAttribute : Attribute
    {
        /// <summary>
        /// Tells the mapper to stop reading data at the specified 1-based row number
        /// </summary>
        /// <param name="i"></param>
        public EndingWithRowAttribute(int i)
        {
        }
    }
}