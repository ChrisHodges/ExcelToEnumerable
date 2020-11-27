using System;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Throws a <see cref="ExcelToEnumerableInvalidHeaderException"/> if a property exists on the mapped class that does not have a corresponding column in the spreadsheet.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class AllPropertiesMustBeMappedToColumnsAttribute : Attribute
    {
        /// <summary>
        /// Pass <c>true</c> to throw an exception or <c>false</c> to ignore unmapped properties.
        /// </summary>
        /// <param name="b"></param>
        public AllPropertiesMustBeMappedToColumnsAttribute(bool b = true)
        {
        }
    }
}