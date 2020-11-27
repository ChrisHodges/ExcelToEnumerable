using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Marks the property as optional. Use when <see cref="AllPropertiesMustBeMappedToColumnsAttribute"/> has been applied to
    /// this property's class.
    /// </summary>
    public class OptionalAttribute : Attribute
    {
    }
}