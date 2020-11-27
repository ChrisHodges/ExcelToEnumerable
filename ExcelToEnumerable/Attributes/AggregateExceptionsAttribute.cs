using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Tells the mapper to read the entire spreadsheet before raising any validation exceptions as an <c>AggregateException</c>
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class AggregateExceptionsAttribute : Attribute
    {
    }
}