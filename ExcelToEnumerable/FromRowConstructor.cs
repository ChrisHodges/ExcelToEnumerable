using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToEnumerable.Exceptions;
using LightWeightExcelReader;

namespace ExcelToEnumerable
{
    internal class FromRowConstructor
    {
        private IDictionary<int, FromCellSetter> _cellSettersDictionaryForRead;
        private object _currentObject;

        public Dictionary<int, object> RowValues { get; set; }
        
        public Dictionary<string, object> RowValuesByColumRef { get; set; } = new Dictionary<string, object>();
        
        public Type Type { get; set; }

        public IEnumerable<FromCellSetter> Setters { get; set; }

        public bool RowIsPopulated { get; set; }

        public void SetCellSettersDictionaryForRead<T>(string[] headerNames, IExcelToEnumerableOptions<T> options)
        {
            if (headerNames == null)
            {
                var i = 0;
                _cellSettersDictionaryForRead = new Dictionary<int, FromCellSetter>();
                foreach (var item in Setters)
                {
                    var columnNumber = options.CustomHeaderNumbers.ContainsKey(item.PropertyName)
                        ? options.CustomHeaderNumbers[item.PropertyName]
                        : i;
                    _cellSettersDictionaryForRead.Add(columnNumber, item);
                    i++;
                }
            } else {

                ValidateColumnNames(headerNames, options);

                var index = 0;
                _cellSettersDictionaryForRead = new Dictionary<int, FromCellSetter>();
                foreach (var headerName in headerNames)
                {
                    var cellSetter = Setters.Where(x => x.ColumnName != null).FirstOrDefault(x => x.ColumnName.ToLowerInvariant() == headerName);
                    if (cellSetter != null)
                    {
                        _cellSettersDictionaryForRead.Add(index, cellSetter);
                    }

                    index++;
                }
            }
        }

        private void ValidateColumnNames<T>(string[] headerNames, IExcelToEnumerableOptions<T> options)
        {
            IEnumerable<string> namesOnSpreadsheet = headerNames;
            if (options.SkippedFields != null)
            {
                namesOnSpreadsheet = namesOnSpreadsheet.Except(options.SkippedFields.Select(y => y.ToLowerInvariant()));
            }

            namesOnSpreadsheet = namesOnSpreadsheet.OrderBy(x => x);
            var namesOnConfig = Setters.Where(x => x.ColumnName != null).Select(x => x.ColumnName.ToLowerInvariant()).OrderBy(y => y);
            if (string.Join(",", namesOnSpreadsheet) != string.Join(",", namesOnConfig))
            {
                throw new ExcelToEnumerableInvalidHeaderException(namesOnConfig.Except(namesOnSpreadsheet), namesOnSpreadsheet.Except(namesOnConfig));
            }
        }

        public void ReadRowValues(SheetReader sheetReader)
        {
            do
            {
                
                if (sheetReader.Address == null)
                {
                    continue;
                }
                var cellRef = new CellRef(sheetReader.Address);
                RowValuesByColumRef.Add(cellRef.Column, sheetReader.Value);
                var colNumber = cellRef.ColumnNumber - 1;
                if (!_cellSettersDictionaryForRead.ContainsKey(colNumber))
                {
                    continue;
                }

                var cellValue = sheetReader.Value;
                if (cellValue != null)
                {
                    RowIsPopulated = true;
                    if (RowValues == null)
                    {
                        RowValues = new Dictionary<int, object>();
                    }

                    RowValues.Add(colNumber, cellValue);
                }
            } while (sheetReader.ReadNextInRow());
        }

        public void Clear()
        {
            if (RowValues != null)
            {
                RowValues.Clear();
                RowValuesByColumRef.Clear();
                RowIsPopulated = false;
            }

            _currentObject = null;
        }

        private bool CheckForMissingFields<T>(IExcelToEnumerableOptions<T> options, List<string> fieldsPresentInRow,
            int rowCount, IList<Exception> exceptionsList)
        {
            var diff = options.LoweredRequiredColumns.Except(fieldsPresentInRow);
            var success = !diff.Any();
            if (!success)
            {
                var firstException = CreateExceptionForMissingField(diff.First(), rowCount);
                switch (options.ExceptionHandlingBehaviour)
                {
                    case ExceptionHandlingBehaviour.ThrowOnFirstException:
                        throw firstException;
                    default:
                        exceptionsList.Add(firstException);
                        foreach (var exception in diff.Skip(1).Select(x => CreateExceptionForMissingField(x, rowCount)))
                        {
                            exceptionsList.Add(exception);
                        }

                        break;
                }
            }
            return success;
        }

        private void HandleException<T>(
            object cellValue,
            string propertyName,
            int rowNumber,
            int columnNumber,
            IExcelToEnumerableOptions<T> options,
            IList<Exception> exceptionsList,
            Exception innerException,
            ExcelToEnumerableValidationCode? validationCode,
            string validationMessage = null)
        {
            var exception = new ExcelToEnumerableCellException(_currentObject, rowNumber, columnNumber, cellValue,
                propertyName, RowValuesByColumRef, validationMessage, innerException, validationCode);
            switch (options.ExceptionHandlingBehaviour)
            {
                case ExceptionHandlingBehaviour.ThrowOnFirstException:
                    throw exception;
                default:
                    exceptionsList.Add(exception);
                    break;
            }
        }

        private object ConvertType(object cellValue, Type type)
        {
            if (type == typeof(string) && !(cellValue is string))
            {
                cellValue = cellValue.ToString();
            }
            if ((type == typeof(decimal) || type == typeof(decimal?)) && cellValue is double)
            {
                cellValue = Convert.ToDecimal(cellValue);
            }
            if ((type == typeof(int?) || type == typeof(int)) && cellValue is double)
            {
                // CSH 14/11/2019 This IF statement tests for a fraction when casting to int
                if (Math.Abs((double)cellValue % 1) > (Double.Epsilon * 100))
                {
                    throw new InvalidCastException();
                }
                cellValue = Convert.ToInt32(cellValue);
            }

            if ((type == typeof(bool?) || type == typeof(bool)) && cellValue is string)
            {
                var normValue = ((string)cellValue).Trim().ToLowerInvariant();
                switch (normValue)
                {
                    case "yes":
                        cellValue = true;
                        break;
                    case "no":
                        cellValue = false;
                        break;
                    case "":
                        cellValue = false;
                        break;
                    default:
                        throw new InvalidCastException();
                }
            }
            return cellValue;
        }

        public bool AddPropertiesFromRowValues<T>(object obj, int rowCount, IExcelToEnumerableOptions<T> options,
            IList<Exception> exceptionsList)
        {
            _currentObject = obj;
            var success = true;
            List<string> fieldsPresentInThisRow = null;
            if (options.RequiredFields != null && options.RequiredFields.Any())
            {
                fieldsPresentInThisRow = new List<string>();
            }

            foreach (var item in RowValues)
            {
                var cellValue = item.Value;
                var fromCellSetter = _cellSettersDictionaryForRead[item.Key];
                var isRequiredAlreadyAdded = false;
                var invalidCast = false;
                if (fromCellSetter.CustomMapping != null)
                {
                    try
                    {
                        cellValue = fromCellSetter.CustomMapping(cellValue);
                        fromCellSetter.Setter(obj, cellValue);
                    }
                    catch (Exception e)
                    {
                        HandleException(cellValue, fromCellSetter.ColumnName, rowCount, item.Key, options, exceptionsList, e, null, "Value is invalid");
                        success = false;
                        invalidCast = true;
                    }
                }
                else
                {
                    try
                    {
                        cellValue = ConvertType(cellValue, fromCellSetter.Type);
                        fromCellSetter.Setter(obj, cellValue);
                    }
                    catch (InvalidCastException e)
                    {
                        HandleException(cellValue, fromCellSetter.ColumnName, rowCount, item.Key, options, exceptionsList, e, null, "Value is invalid");
                        success = false;
                        invalidCast = true;
                    }
                }
                if (fromCellSetter.Validators != null && !invalidCast)
                {
                    foreach (var validator in fromCellSetter.Validators)
                    {
                        if (!validator.Validator(cellValue))
                        {
                            HandleException(cellValue, fromCellSetter.ColumnName, rowCount, item.Key, options, exceptionsList, null,
                                validator.ExcelToEnumerableValidationCode,
                                validator.Message);
                            if (validator.ExcelToEnumerableValidationCode == ExcelToEnumerableValidationCode.Required)
                            {
                                isRequiredAlreadyAdded = true;
                            }
                            success = false;
                            break;
                        }
                    }
                }
                if (fieldsPresentInThisRow != null && !isRequiredAlreadyAdded)
                {
                    fieldsPresentInThisRow.Add(fromCellSetter.ColumnName);
                }
            }

            if (fieldsPresentInThisRow != null)
            {
                var result = CheckForMissingFields(options, fieldsPresentInThisRow, rowCount, exceptionsList);
                if (!result)
                {
                    success = false;
                }
            }

            if (options.RowNumberColumn != null)
            {
                var setter = Setters.First(x => x.PropertyName == options.RowNumberColumn);
                setter.Setter(obj, rowCount + options.StartRow);
            }

            return success;
        }

        private ExcelToEnumerableCellException CreateExceptionForMissingField(string missingField, int rowCount)
        {
            if (!_cellSettersDictionaryForRead.Any(x => x.Value.ColumnName.ToLowerInvariant() == missingField.ToLowerInvariant()))
            {
                throw new Exception($"Unable to map {missingField} to cell setter");
            }
            var setterKvp = _cellSettersDictionaryForRead.First(x => x.Value.ColumnName.ToLowerInvariant() == missingField.ToLowerInvariant());
            var msg =
                "Value is required";
            return new ExcelToEnumerableCellException(_currentObject, rowCount, setterKvp.Key, null, setterKvp.Value.ColumnName, RowValuesByColumRef, msg);
        }

        public void PrepareForRead<T>(string[] headerArray, IExcelToEnumerableOptions<T> options)
        {
            SetCellSettersDictionaryForRead(headerArray, options);
            if (options.Validations != null)
            {
                foreach (var item in options.Validations)
                {
                    var cellSetter = _cellSettersDictionaryForRead.Values.FirstOrDefault(x =>
                        x.ColumnName.ToLowerInvariant() == item.Key.ToLowerInvariant());
                    if (cellSetter != null)
                    {
                        if (cellSetter.Validators == null)
                        {
                            cellSetter.Validators = new List<ExcelCellValidator>();
                        }

                        foreach (var validator in item.Value)
                        {
                            cellSetter.Validators.Add(validator);
                        }
                    }
                }
            }
        }
    }
}