using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Tells the mapper to map this property to the number of the current row, rather than a column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class MapsToRowNumberAttribute : Attribute
    {
    }
}