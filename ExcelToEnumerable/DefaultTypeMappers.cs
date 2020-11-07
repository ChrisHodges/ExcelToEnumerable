using System;
using System.Collections.Generic;

namespace ExcelToEnumerable
{
    internal static class DefaultTypeMappers
    {
        private static readonly Func<object, object> NullableBoolean = o =>
        {
            if (o == null)
            {
                return default(bool?);
            }

            var oStr = o.ToString().ToLowerInvariant();

            switch (oStr)
            {
                case "1":
                case "true":
                case "yes":
                    return true;
                case "0":
                case "false":
                case "no":
                    return false;
                default:
                    throw new InvalidCastException(
                        $"Unable to parse '{(oStr == "" ? "EMPTY STRING" : oStr)}' to type Boolean");
            }
        };

        private static readonly Func<object, object> Boolean = o =>
        {
            if (o == null)
            {
                throw new NullReferenceException();
            }

            return NullableBoolean(o);
        };

        public static Dictionary<Type, Func<object, object>> Dictionary = new Dictionary<Type, Func<object, object>>
        {
            {typeof(bool), Boolean},
            {typeof(bool?), NullableBoolean}
        };
    }
}