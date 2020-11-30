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
    
    /// <summary>
    /// Using this attribute causes cells that are strings, but that contain numeric values, to map to numeric properties. For example,
    /// a string containing the "1,512kg" would map to an integer property with a value of 1512.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class RelaxedNumberMatchingAttribute : Attribute
    {
        /// <summary>
        /// Pass <c>true</c> (default) to enable relaxed number matching or false to disable.
        /// </summary>
        public RelaxedNumberMatchingAttribute(bool b = true)
        {
        }
    }
}