using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Maps from a column at the specified numeric index.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class UsesColumnNumberAttribute : Attribute
    {
        /// <summary>
        /// Pass the 1-based column number you want to map from.
        /// </summary>
        /// <param name="i"></param>
        // ReSharper disable once UnusedParameter.Local
        public UsesColumnNumberAttribute(int i)
        {
        }
    }
}