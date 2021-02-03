using System;
using System.Collections.Generic;
using System.Linq;

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

        public static readonly Dictionary<Type, Func<object, object>> Dictionary = new Dictionary<Type, Func<object, object>>
        {
            {typeof(bool), Boolean},
            {typeof(bool?), NullableBoolean}
        };

        public static Func<object, object> CreateEnumTypeMapper(Type type)
        {
            var baseType = type.GetTypeWithoutNullable();
            var typeIsNullable = baseType != type;
            if (!baseType.IsEnum)
            {
                throw new ArgumentException($"Was expecting an enum type but received '{type}'");
            }
            
            var enumNames = baseType.GetEnumNames().Select(x => x.ToNormalisedVariableName()).ToArray();
            var enumValues = baseType.GetEnumValues().Cast<object>().ToArray();
            var enumInts = enumValues.Cast<int>().ToArray();
            var intLookup = new Dictionary<int, object>();
            var stringLookup = new Dictionary<string, object>();
            for (var i = 0; i < enumNames.Length; i++)
            {
                intLookup.Add(enumInts[i], enumValues[i]);
                stringLookup.Add(enumNames[i], enumValues[i]);
            }

            object ReturnFunc(object o)
            {
                if (o == null)
                {
                    if (typeIsNullable)
                    {
                        return null;
                    }

                    return 0;
                }

                var objectType = o.GetType();
                if (objectType == typeof(int) || objectType == typeof(double))
                {
                    var objectAsInt = objectType == typeof(int) ? (int) o : Convert.ToInt32(o);
                    if (intLookup.ContainsKey(objectAsInt))
                    {
                        return intLookup[objectAsInt];
                    }

                    throw new InvalidCastException($"Expected an int value mapping to a member of '{baseType.Name}' but got '{objectAsInt}'");
                }

                if (objectType == typeof(string))
                {
                    var objectAsString = o.ToString().ToNormalisedVariableName();
                    if (objectAsString == "")
                    {
                        if (typeIsNullable)
                        {
                            return null;
                        }
                        return 0;
                    }
                    if (stringLookup.ContainsKey(objectAsString))
                    {
                        return stringLookup[objectAsString];
                    }
                }

                throw new InvalidCastException();
            }

            return ReturnFunc;
        }
    }
}