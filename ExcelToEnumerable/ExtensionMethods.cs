using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToEnumerable
{
    public static class ExtensionMethods
    {
        private static IExcelToEnumerableOptions<T> BuildOptions<T>(Action<ExcelToEnumerableOptionsBuilder<T>> options)
        {
            var optionsBuilder = new ExcelToEnumerableOptionsBuilder<T>();
            options?.Invoke(optionsBuilder);
            return optionsBuilder.Build();
        }

        public static IEnumerable<T> ExcelToEnumerable<T>(this string excelFilePath,
            Action<IExcelToEnumerableOptionsBuilder<T>> options = null) where T : new()
        {
            var builtOptions = BuildOptions(options);
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelFilePath, ExcelToEnumerableContext.Instance,
                builtOptions);
        }

        public static IEnumerable<T> ExcelToEnumerable<T>(this string excelFilePath,
            IExcelToEnumerableOptionsBuilder<T> options) where T : new()
        {
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelFilePath, ExcelToEnumerableContext.Instance,
                options.Build());
        }


        public static IEnumerable<T> ExcelToEnumerable<T>(this Stream excelStream,
            Action<IExcelToEnumerableOptionsBuilder<T>> options = null) where T : new()
        {
            var builtOptions = BuildOptions(options);
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelStream, ExcelToEnumerableContext.Instance,
                builtOptions);
        }

        public static IEnumerable<T> ExcelToEnumerable<T>(this Stream excelStream,
            IExcelToEnumerableOptionsBuilder<T> options) where T : new()
        {
            var excelToEnumerableMapper = new ExcelToEnumerableMapper<T>();
            return excelToEnumerableMapper.MapExcelToEnumerable(excelStream, ExcelToEnumerableContext.Instance,
                options.Build());
        }
    }
}