using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Prevents the mapper form mapping this property, even if a corresponding column is present in the worksheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class IgnoreColumnAttribute : Attribute
    {
    }
}