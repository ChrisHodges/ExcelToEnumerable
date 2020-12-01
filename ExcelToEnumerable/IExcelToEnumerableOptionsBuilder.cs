using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable
{
    /// <summary>
    /// Fluently creates an ExcelToEnumerableOptions object. This interface is normally accessed via the ExcelToEnumerable <see cref="ExtensionMethods"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IExcelToEnumerableOptionsBuilder<T>
    {
        /// <summary>
        /// Specifies the 1-based row to start reading data from. Defaults to 1.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .StartRow(5)
        /// );
        /// </code>
        /// </example>
        /// <param name="startRow"></param>
        /// <returns>
        /// The builder
        /// </returns>
        IExcelToEnumerableOptionsBuilder<T> StartingFromRow(int startRow);
        
        /// <summary>
        /// Defines whether ExcelToEnumerable maps using column header names (true) or column numbers (false). Defaults to true.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .UsingHeaderNames(false)
        /// );
        /// </code>
        /// </example>
        /// <param name="usingHeaderNames"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> UsingHeaderNames(bool usingHeaderNames);
        
        /// <summary>
        /// Tells the mapper what to do if it encounters a blank spreadsheet row.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .BlankRowBehaviour(BlankRowBehaviour.StopReading)
        /// );
        /// </code>
        /// </example>
        /// <param name="blankRowBehaviour"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> BlankRowBehaviour(BlankRowBehaviour blankRowBehaviour);
        
        /// <summary>
        /// Tells the mapper to aggregate validation exceptions and throw an AggregateException after the whole spreadsheet
        /// has been read.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .AggregateExceptions()
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> AggregateExceptions();
        
        /// <summary>
        /// Instead of throwing validation exceptions, the mapper adds them to the passed IList.
        /// </summary>
        /// <example>
        /// <code>
        /// var myExceptionList = new List&lt;Exception>();
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .OutputExceptionsTo(myExceptionList)
        /// );
        /// </code>
        /// </example>
        /// <param name="list"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> OutputExceptionsTo(IList<Exception> list);

        /// <summary>
        /// The zero-indexed spreadsheet number to read from.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .UsingSheet(3)
        /// );
        /// </code>
        /// </example>
        /// <param name="i"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> UsingSheet(int i);
        
        /// <summary>
        /// The name of the spreadsheet to read from.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .UsingSheet("My Spreadsheet")
        /// );
        /// </code>
        /// </example>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> UsingSheet(string sheetName);

        /// <summary>
        /// Specifies a mapping option for a specific property. See <see cref="IExcelPropertyConfiguration{T,TProperty}"/> for options.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).ShouldBeUnique()
        ///     .Property(y => y.AnotherProperty).MapsToColumnNamed("Column Name")
        /// );
        /// </code>
        /// </example>
        /// <param name="propertyExpression"></param>
        /// <typeparam name="TProperty"></typeparam>
        /// <returns></returns>
        IExcelPropertyConfiguration<T, TProperty> Property<TProperty>(Expression<Func<T, TProperty>> propertyExpression);
        
        /// <summary>
        /// The highest row number to read up to. 1-based. If not set then the mapper will read to the end of the sheet.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .EndingWithRow(200)
        /// );
        /// </code>
        /// </example>
        /// <param name="maxRow"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> EndingWithRow(int maxRow);
        
        /// <summary>
        /// Sets the row number to read the header from.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .HeaderOnRow(3)
        /// );
        /// </code>
        /// </example>
        /// <param name="rowNumber"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> HeaderOnRow(int rowNumber);
        
        /// <summary>
        /// A hook that fires after the header row has been read. 
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .OnReadingHeaderRow((headerDictionary) => {
        ///         foreach (headerEntry in headerDictionary) {
        ///             Console.WriteLine($"Column Number: {item.Key}");
        ///             Console.WriteLine($"Column Header Name: {item.Value}");
        ///         }
        ///     }
        /// );
        /// </code>
        /// </example>
        /// <param name="action">
        /// An action that accepts an columnNumber => columnName dictionary as an argument.
        /// </param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> OnReadingHeaderRow(Action<IDictionary<int, string>> action);

        /// <summary>
        /// Columns without a corresponding property will be ignored instead of throwing an exception.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .AllColumnsMustBeMappedToProperties()
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> AllColumnsMustBeMappedToProperties(bool b);

        /// <summary>
        /// Throws an <see cref="ExcelToEnumerableInvalidHeaderException"/> if there are properties that are not
        /// mapped to a header on the spreadsheet. To override for a specific property use
        /// <c>Property(x => x.MyProperty).Optional()</c>
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .AllPropertiesMustBeMappedToColumns()
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> AllPropertiesMustBeMappedToColumns(bool b);

        /// <summary>
        /// Setting this to true causes cells that are strings, but that contain numeric values, to map to numeric properties. For example,
        /// a string containing the "1,512kg" would map to an integer property with a value of 1512.
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .RelaxedNumberMatching()
        /// );
        /// </code>
        /// </example>
        /// <param name="b"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> RelaxedNumberMatching(bool b);
    }
}