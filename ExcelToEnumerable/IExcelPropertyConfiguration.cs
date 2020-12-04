using System;
using System.Collections.Generic;
using ExcelToEnumerable.Exceptions;

namespace ExcelToEnumerable
{
    /// <summary>
    /// Configures how the mapper should map a specific property.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <typeparam name="TProperty"></typeparam>
    public interface IExcelPropertyConfiguration<T, TProperty>
    {
        /// <summary>
        /// Will throw an exception if the cell value equal to or less than the supplied value.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).ShouldBeGreaterThan(0)
        /// );
        /// </code>
        /// </example>
        /// <param name="minValue"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> ShouldBeGreaterThan(double minValue);
        
        /// <summary>
        /// Will throw an exception if the cell value is equal to or more than the supplied value.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).ShouldBeLessThan(10)
        /// );
        /// </code>
        /// </example>
        /// <param name="maxValue"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> ShouldBeLessThan(double maxValue);
        
        /// <summary>
        /// Will throw an exception if more than one instance of the value appears in the mapped column.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).ShouldBeUnique()
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> ShouldBeUnique();
        
        /// <summary>
        /// Throws an exception if the cell value is null.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).NotNull()
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> NotNull();
        
        /// <summary>
        /// Throws an exception if the cell value is not contained in the given list.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).ShouldBeOneOf(new []{"A", "B", "C"})
        /// );
        /// </code>
        /// </example>
        /// <param name="minValue"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(IEnumerable<TProperty> minValue);
        
        /// <summary>
        /// Throws an exception if the cell value is not contained in the given list.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).ShouldBeOneOf("A", "B", "C", "D")
        /// );
        /// </code>
        /// </example>
        /// <param name="permissableValues"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(params TProperty[] permissableValues);
        
        /// <summary>
        /// Maps an IEnumerable property to a set of one or more columns.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyList).MapFromColumns(new []{"Column1", "Column2", "Column3"})
        /// );
        /// </code>
        /// </example>
        /// <param name="columnNames"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> MapFromColumns(IEnumerable<string> columnNames);
        
        /// <summary>
        /// Maps an IEnumerable property to a set of one or more columns.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyList).MapFromColumns("Column1", "Column2", "Column3")
        /// );
        /// </code>
        /// </example>
        /// <param name="columnNames"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> MapFromColumns(params string[] columnNames);
        
        /// <summary>
        /// Maps the cell value using a the custom mapping function. The mapping function accepts an object as an argument
        /// and returns an object which should be castable to the mapped property.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyBoolean).UsesCustomMapping(m => m?.ToString() == "yes" ? true : false)
        /// );
        /// </code>
        /// </example>
        /// <param name="mapping"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> UsesCustomMapping(Func<object, object> mapping);
        
        /// <summary>
        /// Maps from a column with the specified name.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).MapsToColumnNamed("My Property");
        /// );
        /// </code>
        /// </example>
        /// <param name="columnName"></param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> MapsToColumnNamed(string columnName);
        
        /// <summary>
        /// Does not attempt to map property.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).IgnoreColumn();
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> IgnoreColumn();
        
        /// <summary>
        /// Rather than mapping from a column of the spreadsheet, this property will map to the row number of the spreadsheet.
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.RowNumberProperty).MapsToRowNumber()
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> MapsToRowNumber();
        
        /// <summary>
        /// Maps from a column at the specified index
        /// </summary>
        /// <param name="i"></param>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).MapsToColumnNumber(5)
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> MapsToColumnNumber(int i);

        /// <summary>
        /// Maps from a column with the specified letter, e.g. A, M, AB, ZZA, etc.
        /// </summary>
        /// <param name="columnLetter">
        ///     e.g. A, M, AB, ZZA, etc.
        /// </param>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).MapsToColumnLetter("E")
        /// );
        /// </code>
        /// </example>
        IExcelToEnumerableOptionsBuilder<T> MapsToColumnLetter(string columnLetter);
        
        /// <summary>
        /// The property will be mapped if the column exists, otherwise it will be initialised with its default value.
        /// </summary>
        /// <param name="isOptional"></param>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).OptionalColumn()
        ///     .Property(y => y.MySecondProperty).OptionalColumn(true)
        ///     .Property(y => y.MyThirdProperty).OptionalColumn(false)
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> OptionalColumn(bool isOptional = true);

        /// <summary>
        /// Maps using the default mapper but first validates using the ExcelCellValidator argument
        /// </summary>
        /// <example>
        /// <code>
        /// var result = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .Property(y => y.MyProperty).UsesCustomValidator(o => o != null &amp;&amp; (o.ToString() == "Y" || o.ToString() == "N"))
        /// );
        /// </code>
        /// </example>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> UsesCustomValidator(Func<object, bool> validator, string message);

        /// <summary>
        /// Will throw an <see cref="ExcelToEnumerableInvalidHeaderException"/> if the property does not have a corresponding column in the spreadsheet.
        /// </summary>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T>  RequiredColumn(bool isRequired = true);
    }
}