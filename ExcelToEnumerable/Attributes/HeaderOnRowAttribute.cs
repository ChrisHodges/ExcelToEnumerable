using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Tells the mapper to reader the header names at the specified 1-based row
    /// </summary>
    public class HeaderOnRowAttribute : Attribute
    {
        /// <summary>
        /// Tells the mapper to reader the header names at the specified 1-based row
        /// </summary>
        /// <param name="i"></param>
        public HeaderOnRowAttribute(int i)
        {
        }
    }
}