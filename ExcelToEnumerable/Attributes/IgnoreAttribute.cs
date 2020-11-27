using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Prevents an exception from being thrown if the property does not have a matching column in the spreadsheet. Overrides the
    /// behaviour of the class level <see cref="AllPropertiesMustBeMappedToColumnsAttribute"/>
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class IgnoreAttribute : Attribute
    {
    }
}