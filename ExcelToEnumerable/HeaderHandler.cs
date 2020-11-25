using System.Collections.Generic;
using System.Linq;
using ExcelToEnumerable.Exceptions;
using SpreadsheetCellRef;

namespace ExcelToEnumerable
{
    internal class HeaderHandler
    {
        private void AddCellValidators<T>(IExcelToEnumerableOptions<T> options)
        {
            if (options.Validations == null) return;
            foreach (var item in options.Validations)
            {
                var cellSetter = _cellSettersDictionaryForRead.Values.FirstOrDefault(x =>
                    x.ColumnName?.ToLowerInvariant() == item.Key.ToLowerInvariant());
                if (cellSetter == null) continue;
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
        public IDictionary<int, PropertySetter> GetPropertySetterDictionary<T>(IDictionary<int,string> headers, IExcelToEnumerableOptions<T> options)
        {
            var normalisedHeaderArray = headers?.Values.Select(x => x?.ToNormalisedVariableName()).ToArray();
            GetPropertySetterDictionary(options, normalisedHeaderArray);
            AddCellValidators(options);
            return _cellSettersDictionaryForRead;
        }

        private IEnumerable<PropertySetter> Setters { get; }
        
        private IDictionary<int, PropertySetter> _cellSettersDictionaryForRead;

        public HeaderHandler(IEnumerable<PropertySetter> setters)
        {
            Setters = setters;
        }

        private void GetPropertySetterDictionaryByColumnNumber(Dictionary<string, int> customHeaderNumbers)
        {
            var i = 0;
            _cellSettersDictionaryForRead = new Dictionary<int, PropertySetter>();
            foreach (var item in Setters)
            {
                var columnNumber = customHeaderNumbers.ContainsKey(item.PropertyName)
                    ? customHeaderNumbers[item.PropertyName]
                    : i;
                if (_cellSettersDictionaryForRead.ContainsKey(columnNumber))
                {
                    throw new ExcelToEnumerableConfigException(
                        $"Trying to map property '{item.PropertyName}' to column '{CellRef.NumberToColumnName(columnNumber + 1)}' but that column is already mapped to property '{_cellSettersDictionaryForRead[columnNumber].PropertyName}'. If you're not using header names then all properties need to be mapped to a column or explicitly ignored.");
                }

                _cellSettersDictionaryForRead.Add(columnNumber, item);
                i++;
            }
        }

        private void GetPropertySetterDictionaryByColumnName(string[] normalisedHeaderNames)
        {
            var index = 0;
            _cellSettersDictionaryForRead = new Dictionary<int, PropertySetter>();
            foreach (var headerName in normalisedHeaderNames)
            {
                var cellSetter = Setters.Where(x => x.ColumnName != null)
                    .FirstOrDefault(x => x.ColumnName.ToNormalisedVariableName() == headerName);
                if (cellSetter != null)
                {
                    _cellSettersDictionaryForRead.Add(index, cellSetter);
                }

                index++;
            }
        }
        
        private IEnumerable<string> GetNormalisedColumnNames(IEnumerable<PropertySetter> setters)
        {
            return setters.Where(x => x.ColumnName != null).Select(x => x.ColumnName.ToNormalisedVariableName())
                .OrderBy(y => y);
        }

        private static void ValidateColumnNames<T>(IExcelToEnumerableOptions<T> options,
            IEnumerable<string> normalisedColumnNames,
            IEnumerable<string> normalisedHeaderArray)
        {
            IEnumerable<string> namesOnSpreadsheet = normalisedHeaderArray;

            //Exclude unmapped properties
            if (options.UnmappedProperties != null)
            {
                namesOnSpreadsheet =
                    namesOnSpreadsheet.Except(options.UnmappedProperties.Select(y => y.ToLowerInvariant()));
            }

            //Exclude optional properties
            if (options.OptionalProperties != null)
            {
                namesOnSpreadsheet =
                    namesOnSpreadsheet.Except(options.OptionalProperties.Select(y => y.ToLowerInvariant()))
                        .OrderBy(y => y);
            }
            
            namesOnSpreadsheet = namesOnSpreadsheet.OrderBy(x => x);

            var namesOnConfig = normalisedColumnNames.ToArray();
            if (options.OptionalProperties != null)
            {
                namesOnConfig =
                    namesOnConfig.Except(options.OptionalProperties.Select(y => y.ToLowerInvariant())).OrderBy(y => y)
                        .ToArray();
            }

            if (options.IgnoreColumnsWithoutMatchingProperties)
            {
                if (namesOnConfig.Except(namesOnSpreadsheet).Any())
                {
                    throw new ExcelToEnumerableInvalidHeaderException(namesOnConfig.Except(namesOnSpreadsheet),
                        namesOnConfig.Except(namesOnConfig));
                }
            }
            else
            {
                if (string.Join(",", namesOnSpreadsheet) != string.Join(",", namesOnConfig))
                {
                    throw new ExcelToEnumerableInvalidHeaderException(namesOnConfig.Except(namesOnSpreadsheet),
                        namesOnSpreadsheet.Except(namesOnConfig));
                }
            }
        }
        
        private void GetPropertySetterDictionary<T>(IExcelToEnumerableOptions<T> options,
            string[] normalisedHeaderArray)
        {
            if (normalisedHeaderArray == null)
            {
                GetPropertySetterDictionaryByColumnNumber(options.CustomHeaderNumbers);
            }
            else
            {
                ValidateColumnNames(options, GetNormalisedColumnNames(Setters), normalisedHeaderArray);
                GetPropertySetterDictionaryByColumnName(normalisedHeaderArray);
            }
        }
    }
}