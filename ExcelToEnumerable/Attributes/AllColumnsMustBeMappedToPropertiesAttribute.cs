using System;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Throws a <see cref="ExcelToEnumerableInvalidHeaderException"/> if a column exists on the spreadsheet that does not have a corresponding property in the mapped class.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class AllColumnsMustBeMappedToPropertiesAttribute : Attribute
    {
        /// <summary>
        /// Pass <c>true</c> (default) to throw an exception or false to ignore unmapped columns.
        /// </summary>
        /// <param name="b"></param>
        // ReSharper disable once UnusedParameter.Local
        public AllColumnsMustBeMappedToPropertiesAttribute(bool b = true)
        {
        }
    }
}