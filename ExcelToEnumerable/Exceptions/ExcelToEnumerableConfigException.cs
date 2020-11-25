using System;

namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Thrown when an invalid configuration is encountered by the mapper.
    /// </summary>
    public class ExcelToEnumerableConfigException : ExcelToEnumerableException
    {
        internal ExcelToEnumerableConfigException(string message) : base(message)
        {
        }
    }
}