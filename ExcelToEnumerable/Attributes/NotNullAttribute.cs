using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Throws a validation exception if the cell value is null.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class NotNullAttribute : Attribute
    {
    }
}