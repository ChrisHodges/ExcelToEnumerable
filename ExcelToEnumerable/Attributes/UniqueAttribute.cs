using System;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Will throw an <see cref="ExcelToEnumerableSheetException"/> if the property is duplicated within the spreadsheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class UniqueAttribute : Attribute
    {
    }
}