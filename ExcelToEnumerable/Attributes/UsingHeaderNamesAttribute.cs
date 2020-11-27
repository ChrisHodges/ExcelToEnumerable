using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Maps properties by header name if true. Maps by column number if false
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class UsingHeaderNamesAttribute : Attribute
    {
        /// <summary>
        /// Pass <c>true</c> (default) to throw map by header name or false to map by column number.
        /// </summary>
        /// <param name="b"></param>
        public UsingHeaderNamesAttribute(bool b)
        {
        }
    }
}