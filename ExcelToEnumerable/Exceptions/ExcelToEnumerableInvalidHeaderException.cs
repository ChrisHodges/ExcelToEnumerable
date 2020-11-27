using System.Collections.Generic;
using System.Linq;

namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Thrown when the properties are not correctly mapped to headers. By default, ExcelToEnumerable expects all
    /// properties to have one or more corresponding headers. To ignore a property, use:
    /// <c>
    /// optionsBuilder.Property(y => y.IgnoreThisProperty).Ignore()
    /// </c>
    /// To ignore columns without matching properties use:
    /// <c>
    /// optionsBuilder.AllColumnsMustBeMappedToProperties()
    /// </c>
    /// </summary>
    public class ExcelToEnumerableInvalidHeaderException : ExcelToEnumerableSheetException
    {
        internal ExcelToEnumerableInvalidHeaderException(IEnumerable<string> missingHeaders,
            IEnumerable<string> missingProperties) : base(
            BuildExceptionMessage(missingHeaders, missingProperties)
        )
        {
            MissingHeaders = missingHeaders;
            MissingProperties = missingProperties;
        }

        /// <summary>
        /// Column headers that were expected to be present in the worksheet, but were not.
        /// </summary>
        public IEnumerable<string> MissingHeaders { get; }
        
        /// <summary>
        /// Properties that were expected to exist on the mapped class, but were not
        /// </summary>
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