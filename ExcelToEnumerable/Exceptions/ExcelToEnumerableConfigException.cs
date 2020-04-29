using System;

namespace ExcelToEnumerable.Exceptions
{
    public class ExcelToEnumerableConfigException : ExcelToEnumerableException
    {
        public ExcelToEnumerableConfigException(string message, Exception exception) : base(message, exception)
        {
        }

        public ExcelToEnumerableConfigException(string message) : base(message)
        {
        }
    }
}