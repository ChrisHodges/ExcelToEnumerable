using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Maps from a column at the specific letter index, (i.e. "A", "C", "AA", etc)
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class MapsToColumnLetterAttribute : Attribute
    {
        /// <summary>
        /// Pass the 1-based column number you want to map from.
        /// </summary>
        /// <param name="i"></param>
        public MapsToColumnLetterAttribute(string i)
        {
        }
    }
}