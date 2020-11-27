using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Maps from the column with the specified header name
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class UsesColumnNamedAttribute : Attribute
    {
        /// <summary>
        /// Maps from the column with the specified header name
        /// </summary>
        /// <param name="columnName"></param>
        // ReSharper disable once UnusedParameter.Local
        public UsesColumnNamedAttribute(string columnName)
        {
        }
    }
}
