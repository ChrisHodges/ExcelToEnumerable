using System;

namespace ExcelToEnumerable
{
    public static class RowAndColumnExtensionMethods
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="i">
        /// This is 1-based not 0-based</param>
        /// <returns></returns>
        public static string ToColumnName(this int i)
        {
            var dividend = i;
            var columnName = "";
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = Convert.ToInt32((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}