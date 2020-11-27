using System;

namespace ExcelToEnumerable.Attributes
{
    /// <summary>
    /// Defines how the mapper should behave when encountering a blank row
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class WithBlankRowBehaviourAttribute : Attribute
    {
        /// <summary>
        /// The mapper will behave in the way specified by the specified <see cref="BlankRowBehaviour"/>
        /// </summary>
        /// <param name="blankRowBehaviour"></param>
        public WithBlankRowBehaviourAttribute(BlankRowBehaviour blankRowBehaviour)
        {
        }
    }
}