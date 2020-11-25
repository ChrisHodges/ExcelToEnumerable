using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToEnumerable
{
    /// <summary>
    /// Extension methods to map an Excel worksheet to an enumerable of type T
    /// </summary>
    public static class ExtensionMethods
    {
        private static IExcelToEnumerableOptions<T> BuildOptions<T>(Action<ExcelToEnumerableOptionsBuilder<T>> options)
        {
            var optionsBuilder = new ExcelToEnumerableOptionsBuilder<T>();
            options?.Invoke(optionsBuilder);
            return optionsBuilder.Build();
        }

        /// <summary>
        /// Maps the spreadsheet at the given filepath to an enumerable of type T, using an optional fluent options expression
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = "C://path/to/spreadsheet".ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .UsingSheet("My Spreadsheet")
        /// );
        /// </code>
        /// </example>
        /// <param name="excelFilePath">Path to an Excel spreadsheet</param>
        /// <param name="options">a fluent <see cref="IExcelToEnumerableOptionsBuilder{T}"/> expression</param>
        /// <typeparam name="T"></typeparam>
        /// <returns>An <c>IEnumerable</c> of type <c>T</c></returns>
        public static IEnumerable<T> ExcelToEnumerable<T>(this string excelFilePath,
            Action<IExcelToEnumerableOptionsBuilder<T>> options = null) where T : new()
        {
            var builtOptions = BuildOptions(options);
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelFilePath, ExcelToEnumerableContext.Instance,
                builtOptions);
        }

        /// <summary>
        /// Maps the spreadsheet at the given path to an enumerable of type T using an options builder
        /// </summary>
        /// <example>
        /// <code>
        /// var optionsBuilder = new ExcelToEnumerableOptionsBuilder&lt;TestClass&gt;();
        /// optionsBuilder.AggregateExceptions();
        /// IEnumerable&lt;MyClass&gt; results = "C://path/to/spreadsheet".ExcelToEnumerable&lt;MyClass&gt;(optionsBuilder);
        /// </code>
        /// </example>
        /// <param name="excelFilePath"></param>
        /// <param name="options"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static IEnumerable<T> ExcelToEnumerable<T>(this string excelFilePath,
            IExcelToEnumerableOptionsBuilder<T> options) where T : new()
        {
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelFilePath, ExcelToEnumerableContext.Instance,
                ((ExcelToEnumerableOptionsBuilder<T>)options).Build());
        }


        /// <summary>
        /// Maps the worksheet Stream to an enumerable of type T, using an optional fluent options expression
        /// </summary>
        /// <example>
        /// <code>
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(x => x
        ///     .UsingSheet("My Spreadsheet")
        /// );
        /// </code>
        /// </example>
        /// <param name="excelStream"></param>
        /// <param name="options"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static IEnumerable<T> ExcelToEnumerable<T>(this Stream excelStream,
            Action<IExcelToEnumerableOptionsBuilder<T>> options = null) where T : new()
        {
            var builtOptions = BuildOptions(options);
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelStream, ExcelToEnumerableContext.Instance,
                builtOptions);
        }

        /// <summary>
        /// Maps the worksheet Stream to an enumerable of type T using options in the options builder
        /// </summary>
        /// <example>
        /// <code>
        /// var optionsBuilder = new ExcelToEnumerableOptionsBuilder&lt;TestClass&gt;();
        /// optionsBuilder.AggregateExceptions();
        /// IEnumerable&lt;MyClass&gt; results = excelStream.ExcelToEnumerable&lt;MyClass&gt;(optionsBuilder);
        /// </code>
        /// </example>
        /// <param name="excelStream"></param>
        /// <param name="options"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static IEnumerable<T> ExcelToEnumerable<T>(this Stream excelStream,
            IExcelToEnumerableOptionsBuilder<T> options) where T : new()
        {
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelStream, ExcelToEnumerableContext.Instance,
                ((ExcelToEnumerableOptionsBuilder<T>)options).Build());
        }
    }
}