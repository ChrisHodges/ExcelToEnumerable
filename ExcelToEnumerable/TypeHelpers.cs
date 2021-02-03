using System;

namespace ExcelToEnumerable
{
    internal static class TypeHelpers
    {
        public static bool IsNumeric(this Type type)
        {
            var baseType = type.GetTypeWithoutNullable();
            return baseType == typeof(int) || baseType == typeof(double) || baseType == typeof(decimal) || baseType == typeof(float) || baseType == typeof(long);
        }
        
        public static Type GetTypeWithoutNullable(this Type type)
        {
            var baseType = Nullable.GetUnderlyingType(type);
            if (baseType == null)
            {
                baseType = type;
            }
            return baseType;
        }
    }
}