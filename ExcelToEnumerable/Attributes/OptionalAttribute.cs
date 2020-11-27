using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Marks the property as optional. Use when <see cref="AllPropertiesMustBeMappedToColumnsAttribute"/> has been applied to
    /// this property's class.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class OptionalAttribute : Attribute
    {
    }
    
    /// <summary>
    /// Marks the property as required.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class RequiredAttribute : Attribute
    {
    }
}