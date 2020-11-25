namespace ExcelToEnumerable.Exceptions
{
    /// <summary>
    /// Enum. Enumerates reasons why a cell my fail validation. This is not a exhaustive list since a property can have custom validations.
    /// </summary>
    public enum ExcelToEnumerableValidationCode
    {
        /// <summary>
        /// The property value cannot be null.
        /// </summary>
        Required,
        
        /// <summary>
        /// The property value needs to be one of a discrete list of values.
        /// </summary>
        OneOf,
        
        /// <summary>
        /// The property value must be less than a specified number.
        /// </summary>
        LessThan,
        
        /// <summary>
        /// The property value must be greater than a specified number.
        /// </summary>
        GreaterThan
    }
}