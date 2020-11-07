using System.Collections.Generic;
using System.Linq;

namespace ExcelToEnumerable.Exceptions
{
    public class ExcelToEnumerableInvalidHeaderException : ExcelToEnumerableSheetException
    {
        public ExcelToEnumerableInvalidHeaderException(IEnumerable<string> missingHeaders,
            IEnumerable<string> missingProperties) : base(
            BuildExceptionMessage(missingHeaders, missingProperties)
        )
        {
            MissingHeaders = missingHeaders;
            MissingProperties = missingProperties;
        }

        public IEnumerable<string> MissingHeaders { get; }
        public IEnumerable<string> MissingProperties { get; }

        private static string BuildExceptionMessage(IEnumerable<string> missingHeaders,
            IEnumerable<string> missingProperties)
        {
            var missingHeadersMessage = missingHeaders != null && missingHeaders.Any()
                ? $"Missing headers: {string.Join(", ", missingHeaders.Select(x => $"'{x}'"))}. "
                : "";
            var missingPropertyMessages = missingProperties != null && missingProperties.Any()
                ? $"Missing properties: {string.Join(", ", missingProperties.Select(x => $"'{x}'"))}."
                : "";
            return $"{missingHeadersMessage}{missingPropertyMessages}";
        }
    }
}