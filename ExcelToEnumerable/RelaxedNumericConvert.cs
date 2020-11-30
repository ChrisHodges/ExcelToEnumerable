using System;
using System.Text.RegularExpressions;

namespace ExcelToEnumerable
{
    internal static class RelaxedNumericConvert
    {
        private static Regex _numberMatcher = new Regex(@"([-0123456789,.]+)");
        
        public static double ToDouble(string inputStr)
        {
            var match = _numberMatcher.Match(inputStr);
            if (match.Success)
            {
                return Convert.ToDouble(match.Groups[1].Value);
            }

            throw new FormatException();
        }
    }
}