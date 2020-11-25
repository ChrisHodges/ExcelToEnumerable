using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using ExcelToEnumerable.Exceptions;
using LightweightExcelReader;
using SpreadsheetCellRef;

[assembly: InternalsVisibleTo("ExcelToEnumerable.Tests")]

namespace ExcelToEnumerable
{
    internal class ExcelToEnumerableMapper<T> where T : new()
    {
        private List<Exception> _exceptionList;

        private IExcelToEnumerableOptions<T> _options;

        private IEnumerable<T> HandleCreateEntityBehaviour(SheetReader worksheet, RowMapper rowMapper)
        {
            var list = new List<T>();
            var rowDiff = worksheet.CurrentRowNumber - worksheet.PreviousRowNumber;
            if (rowDiff > 1)
            {
                for (var i = 1; i < rowDiff; i++)
                {
                    list.Add(new T());
                }
            }

            var obj = new T();
            if (rowMapper.RowIsPopulated)
            {
                rowMapper.AddPropertiesFromRowValues(obj, worksheet.CurrentRowNumber.Value, _options,
                    _exceptionList);
            }

            list.Add(obj);
            return list;
        }

        private IEnumerable<T> MainLoop(SheetReader worksheet, RowMapper rowMapper,
            IExcelToEnumerableOptions<T> options, IDictionary<int, PropertySetter> cellSettersDictionaryForRead)
        {
            rowMapper.SetCellSettersDictionaryForRead(cellSettersDictionaryForRead);
            var list = new List<T>();
            do
            {
                if (worksheet.CurrentRowNumber > _options.EndRow)
                {
                    break;
                }

                rowMapper.ReadRowValues(worksheet);
                var rowNumber = worksheet.CurrentRowNumber.Value;
                switch (options.BlankRowBehaviour)
                {
                    case BlankRowBehaviour.CreateEntity:
                        var newEntities = HandleCreateEntityBehaviour(worksheet, rowMapper);
                        list.AddRange(newEntities);
                        break;
                    case BlankRowBehaviour.Ignore:
                        var item = HandleIgnoreBehaviour(rowMapper, rowNumber);
                        if (item != null)
                        {
                            list.Add(item);
                        }

                        break;
                    case BlankRowBehaviour.StopReading:
                        var continu =
                            HandleStopReadingOrThrowExceptionBehaviour(worksheet, list, rowMapper);
                        if (!continu)
                        {
                            return list;
                        }

                        break;
                    case BlankRowBehaviour.ThrowException:
                        var addedRow =
                            HandleStopReadingOrThrowExceptionBehaviour(worksheet, list, rowMapper);
                        if (!addedRow)
                        {
                            throw new ExcelToEnumerableRowException(
                                null,
                                $"Blank row found at row: '{rowNumber}'", rowNumber,
                                rowMapper.RowValuesByColumRef);
                        }

                        break;
                }

                rowMapper.Clear();
            } while (worksheet.ReadNext());

            return list;
        }

        private IEnumerable<T> MapWorkbookToEnumerable(ExcelReader workbook,
            IExcelToEnumerableContext excelToEnumerableContext,
            IExcelToEnumerableOptions<T> options)
        {
            _options = options;
            var worksheet = _options.WorksheetNumber.HasValue
                ? workbook[_options.WorksheetNumber.Value]
                : workbook[_options.WorksheetName];

            while (!worksheet.CurrentRowNumber.HasValue || worksheet.CurrentRowNumber < _options.HeaderRow)
            {
                worksheet.ReadNext();
            }
            
            _exceptionList =
                options.ExceptionHandlingBehaviour == ExceptionHandlingBehaviour.ThrowOnFirstException
                    ? null
                    : new List<Exception>();

            var rowMapper = excelToEnumerableContext.GetRowMapper(_options) ??
                                     excelToEnumerableContext.SetRowMapper(_options);

            var headerArray = ReadHeader(options, worksheet);
            var headerHandler = new HeaderHandler(rowMapper.Setters);
            var cellSettersDictionaryForRead = headerHandler.GetPropertySetterDictionary(headerArray, options);

            var list = MainLoop(worksheet, rowMapper, options, cellSettersDictionaryForRead);

            if (_options.UniqueProperties != null)
            {
                list = CheckUniqueFields(rowMapper, list);
            }

            HandleAggregatedExceptions();

            return list;
        }

        private IEnumerable<T> CheckUniqueFields(RowMapper rowMapper, IEnumerable<T> list)
        {
            var itemsToRemove = new List<T>();
            foreach (var uniqueField in _options.UniqueProperties)
            {
                var getter = rowMapper.Setters.First(x => x.ColumnName == uniqueField.ToLowerInvariant())
                    .Getter;
                var kvps = list.Select(x => new KeyValuePair<object, T>(getter(x), x));
                var dupes = kvps.GroupBy(x => x.Key)
                    .Where(g => g.Count() > 1)
                    .ToList();
                if (dupes.Any())
                {
                    itemsToRemove.AddRange(dupes.SelectMany(x => x.Select(y => y.Value)));
                    var exception = new ExcelToEnumerableSheetException(
                        $"Duplicate values for column '{uniqueField}': {string.Join(", ", dupes.Select(x => x.Key))}");

                    //CSH 2019-11-12 If the exception list is not null we're either throwing an aggregate exeption
                    //or returning an exception list. Either way, we don't want to throw the exception now.
                    if (_exceptionList != null)
                    {
                        _exceptionList.Add(exception);
                    }
                    else
                    {
                        throw exception;
                    }
                }
            }

            list = list.Where(x => !itemsToRemove.Contains(x)).ToList();
            return list;
        }

        internal IEnumerable<T> MapExcelToEnumerable(string filePath,
            IExcelToEnumerableContext excelToEnumerableContext,
            IExcelToEnumerableOptions<T> options)
        {
            var workbook = new ExcelReader(filePath);
            var list = MapWorkbookToEnumerable(workbook, excelToEnumerableContext, options);
            return list;
        }

        internal IEnumerable<T> MapExcelToEnumerable(Stream excelStream,
            ExcelToEnumerableContext excelToEnumerableContext, IExcelToEnumerableOptions<T> options)
        {
            var workbook = new ExcelReader(excelStream);
            var list = MapWorkbookToEnumerable(workbook, excelToEnumerableContext, options);
            excelStream.Dispose();
            return list;
        }

        private void HandleAggregatedExceptions()
        {
            if (_exceptionList != null && _exceptionList.Any())
            {
                if (_options.ExceptionHandlingBehaviour == ExceptionHandlingBehaviour.AggregateExceptions)
                {
                    throw new AggregateException(_exceptionList);
                }

                if (_options.ExceptionHandlingBehaviour == ExceptionHandlingBehaviour.LogExceptions)
                {
                    foreach (var exception in _exceptionList)
                    {
                        _options.ExceptionList.Add(exception);
                    }
                }
            }
        }

        private IDictionary<int, string> ReadHeader(IExcelToEnumerableOptions<T> options, SheetReader worksheet)
        {
            IDictionary<int, string> header = null;
            if (options.UseHeaderNames || _options.OnReadingHeaderRowAction != null)
            {
                header = GetHeaderRow(worksheet);
                _options.OnReadingHeaderRowAction?.Invoke(header);
            }

            while (worksheet.CurrentRowNumber < options.StartRow)
            {
                worksheet.ReadNext();
            }

            return header;
        }

        private bool HandleStopReadingOrThrowExceptionBehaviour(SheetReader worksheet, List<T> list,
            RowMapper rowMapper)
        {
            var rowDiff = worksheet.CurrentRowNumber - worksheet.PreviousRowNumber;
            if (rowDiff > 1)
            {
                return false;
            }

            if (rowMapper.RowIsPopulated)
            {
                var obj = new T();
                if (rowMapper.AddPropertiesFromRowValues(obj, worksheet.CurrentRowNumber.Value, _options,
                    _exceptionList))
                {
                    list.Add(obj);
                }

                return true;
            }

            return false;
        }

        private T HandleIgnoreBehaviour(RowMapper rowMapper, int rowCount)
        {
            if (rowMapper.RowIsPopulated)
            {
                var obj = new T();
                if (rowMapper.AddPropertiesFromRowValues(obj, rowCount, _options, _exceptionList))
                {
                    return obj;
                }
            }

            return default;
        }

        private static Dictionary<int, string> GetHeaderRow(SheetReader worksheet)
        {
            var dict = new Dictionary<int, string>();
            var previousColumnNumber = 0;
            do
            {
                var cellRef = new CellRef(worksheet.Address);
                var columnNumber = cellRef.ColumnNumber;
                for (var i = previousColumnNumber + 1; i < columnNumber; i++)
                {
                    dict.Add(i, $"**blank column** ({CellRef.NumberToColumnName(i)})");
                    //dict.Add(i, null);
                }

                dict.Add(cellRef.ColumnNumber, worksheet.Value.ToString());
                previousColumnNumber = columnNumber;
            } while (worksheet.ReadNextInRow());

            return dict;
        }
    }
}