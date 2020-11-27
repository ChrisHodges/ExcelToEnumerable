using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Tells the mapper to starting reading data from the specified 1-based row number
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class StartingFromRowAttribute : Attribute
    {
        /// <summary>
        /// Tells the mapper to starting reading data from the specified 1-based row number
        /// </summary>
        /// <param name="i"></param>
        public StartingFromRowAttribute(int i)
        {
        }
    }
}