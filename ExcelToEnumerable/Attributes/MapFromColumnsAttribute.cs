using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Use with <c>IEnumerable</c> properties. Maps enumerable property to specific columns.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class MapFromColumnsAttribute : Attribute
    {
        /// <summary>
        /// Maps the <c>IEnumerable</c> to the given columns
        /// </summary>
        /// <param name="columns"></param>
        /// <exception cref="NotImplementedException"></exception>
        // ReSharper disable once UnusedParameter.Local
        public MapFromColumnsAttribute(params string[] columns)
        {
        }
    }
}