namespace ExcelToEnumerable
{
    internal static class VariableNameNormalisationHelper
    {
        public static string ToNormalisedVariableName(this string str)
        {
            return str?.Trim().ToLowerInvariant().Replace(" ", "").Replace("-", "").Replace("_", "");
        }
    }
}