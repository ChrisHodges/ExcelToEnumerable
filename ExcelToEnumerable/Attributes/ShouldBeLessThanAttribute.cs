using System;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Throws an <see cref="ExcelToEnumerableCellException"/> if the cell value is equal to or greater than the specified value
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ShouldBeLessThanAttribute : Attribute
    {
        /// <summary>
        /// Throws an <see cref="ExcelToEnumerableCellException"/> if the cell value is equal to or greater than the specified value
        /// </summary>
        /// <param name="i"></param>
        public ShouldBeLessThanAttribute(double i)
        {
        }
    }
    
    /// <summary>
    /// Throws an <see cref="ExcelToEnumerableCellException"/> if the cell value is equal to or less than the specified value
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ShouldBeGreaterThanAttribute : Attribute
    {
        /// <summary>
        /// Throws an <see cref="ExcelToEnumerableCellException"/> if the cell value is equal to or less than the specified value
        /// </summary>
        /// <param name="i"></param>
        public ShouldBeGreaterThanAttribute(double i)
        {
        }
    }
}