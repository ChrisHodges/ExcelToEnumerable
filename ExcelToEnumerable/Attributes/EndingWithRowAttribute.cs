using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Tells the mapper to stop reading data at the specified 1-based row number
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