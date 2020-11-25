using System;

namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Abstract base class from which all ExcelToEnumerable exceptions derive.
    /// </summary>
    public abstract class ExcelToEnumerableException : Exception
    {
        internal ExcelToEnumerableException(string message, Exception exception) : base(message, exception)
        {
        }

        internal ExcelToEnumerableException(string message) : base(message)
        {
        }
    }
}