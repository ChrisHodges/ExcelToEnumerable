using System.Collections.Generic;

namespace ExcelToEnumerable.Exceptions
{
    public class ExcelToEnumerableInvalidHeaderException : ExcelToEnumerableSheetException
    {
        public ExcelToEnumerableInvalidHeaderException(IEnumerable<string> missingHeaders, IEnumerable<string> missingProperties) : base(
            $"Missing headers: {string.Join(", ", missingHeaders)}. Missing properties: {string.Join(", ", missingProperties)}"
            )
        {
            MissingHeaders = missingHeaders;
            MissingProperties = missingProperties;
        }

        public IEnumerable<string> MissingHeaders { get; }
        public IEnumerable<string> MissingProperties { get; }
    }
}