using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable
{
    internal static class ExcelCellValidatorFactory
    {
        internal static ExcelCellValidator CreateGreaterThan(double minValue)
        {
            return new ExcelCellValidator
            {
                Message = $"Should be greater than {minValue}",
                Validator = o => !(o is string) && Convert.ToDouble(o) > minValue,
                ExcelToEnumerableValidationCode = ExcelToEnumerableValidationCode.GreaterThan
            };
        }

        internal static ExcelCellValidator CreateLessThan(double maxValue)
        {
            return new ExcelCellValidator
            {
                Message = $"Should be less than {maxValue}",
                Validator = o => !(o is string) && Convert.ToDouble(o) < maxValue,
                ExcelToEnumerableValidationCode = ExcelToEnumerableValidationCode.LessThan
            };
        }

        internal static ExcelCellValidator CreateShouldBeOneOf<TProperty>(IEnumerable<TProperty> oneOfArray)
        {
            return new ExcelCellValidator
            {
                Message = $"Should be one of {string.Join(", ", oneOfArray.Select(y => y.ToString()))}",
                Validator = o => oneOfArray.Contains((TProperty) o),
                ExcelToEnumerableValidationCode = ExcelToEnumerableValidationCode.OneOf
            };
        }

        internal static ExcelCellValidator CreateRequired()
        {
            return new ExcelCellValidator
            {
                Message = "Value is required",
                Validator = o =>
                {
                    if (o == null)
                    {
                        return false;
                    }

                    if (o is string s && s == "")
                    {
                        return false;
                    }

                    return true;
                },
                ExcelToEnumerableValidationCode = ExcelToEnumerableValidationCode.Required
            };
        }
    }
}