using System;

namespace ExcelToEnumerable.Exceptions
{
    public abstract class ExcelToEnumerableException : Exception
    {
        protected ExcelToEnumerableException(string message, Exception exception) : base(message, exception)
        {
        }

        protected ExcelToEnumerableException(string message) : base(message)
        {
        }
    }
}