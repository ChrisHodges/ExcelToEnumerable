using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Throws an exception if the mapped value is not one of the specified values.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ShouldBeOneOfAttribute : Attribute
    {
        /// <summary>
        /// Throws an exception if the mapped value is not one of the specified values.
        /// </summary>
        /// <param name="strings"></param>
        /// <exception cref="NotImplementedException"></exception>
        public ShouldBeOneOfAttribute(params object[] strings)
        {
            throw new NotImplementedException();
        }
    }
}