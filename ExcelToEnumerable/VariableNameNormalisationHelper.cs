using System.Text.RegularExpressions;

namespace ExcelToEnumerable
{
    internal static class VariableNameNormalisationHelper
    {
        private static Regex _invalidCharacters = new Regex("[^a-zA-Z0-9@]");
        public static string ToNormalisedVariableName(this string str)
        {
            return _invalidCharacters.Replace(str?.ToLowerInvariant(), "");
        }
    }
}