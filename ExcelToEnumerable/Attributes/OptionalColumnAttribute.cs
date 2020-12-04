using System;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Marks the property as optional. Use when <see cref="AllPropertiesMustBeMappedToColumnsAttribute"/> has been applied to
    /// this property's class.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class OptionalColumnAttribute : Attribute
    {
    }
    
    /// <summary>
    /// Will throw an <see cref="ExcelToEnumerableInvalidHeaderException"/> if the property does not have a corresponding column in the spreadsheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class RequiredColumnAttribute : Attribute
    {
    }
}