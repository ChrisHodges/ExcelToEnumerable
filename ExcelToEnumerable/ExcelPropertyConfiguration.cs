using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToEnumerable.Exceptions;
using SpreadsheetCellRef;

namespace ExcelToEnumerable
{
    internal static class ExcelPropertyConfiguration
    {
        internal static void UsesColumnNumber<T>(int i, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (i < 1)
            {
                throw new ExcelToEnumerableConfigException(
                    $"Unable to map '{propertyName}' to column {i}. UsesColumnNumber expects a 1-based column number");
            }

            options.CustomHeaderNumbers[propertyName] = i - 1;
        }
        
        internal static void Optional<T>(bool isOptional, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (isOptional)
            {
                options.OptionalProperties.Add(propertyName);
                options.ExplicitlyRequiredProperties.Remove(propertyName);
            }
            else
            {
                options.OptionalProperties.Remove(propertyName);
                options.ExplicitlyRequiredProperties.Add(propertyName);
            }
        }

        internal static void MapFromColumns<T>(IEnumerable<string> columnNames, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (options.CollectionConfigurations == null)
            {
                options.CollectionConfigurations = new Dictionary<string, ExcelToEnumerableCollectionConfiguration>();
            }

            var configuration = new ExcelToEnumerableCollectionConfiguration
            {
                PropertyName = propertyName,
                ColumnNames = columnNames
            };

            options.CollectionConfigurations.Add(propertyName, configuration);
        }
    }
    
    internal class ExcelPropertyConfiguration<T, TProperty> : IExcelPropertyConfiguration<T, TProperty>
    {
        private readonly IExcelToEnumerableOptions<T> _options;
        private readonly string _propertyName;
        private readonly IExcelToEnumerableOptionsBuilder<T> _optionsBuilder;

        public ExcelPropertyConfiguration(IExcelToEnumerableOptionsBuilder<T> optionsBuilder, string propertyName,
            IExcelToEnumerableOptions<T> options)
        {
            _options = options;
            _optionsBuilder = optionsBuilder;
            _propertyName = propertyName;
            CreateDictionaryIfNotExists();
            CreateEntryIfNotExists();
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeGreaterThan(double minValue)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateGreaterThan(minValue));
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeLessThan(double maxValue)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateLessThan(maxValue));
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(IEnumerable<TProperty> oneOf)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateShouldBeOneOf(oneOf));
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(params TProperty[] permissiableValues)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateShouldBeOneOf(permissiableValues));
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapFromColumns(IEnumerable<string> columnNames)
        {
            ExcelPropertyConfiguration.MapFromColumns(columnNames, _propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapFromColumns(params string[] columnNames)
        {
            return MapFromColumns(columnNames.ToList());
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeUnique()
        {
            if (_options.UniqueProperties == null)
            {
                _options.UniqueProperties = new List<string>();
            }

            _options.UniqueProperties.Add(_propertyName);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> IsRequired()
        {
            _options.RequiredFields.Add(_propertyName);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsesCustomMapping(Func<object, object> mapping)
        {
            _options.CustomMappings[_propertyName] = mapping;
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsesColumnNamed(string columnName)
        {
            _options.CustomHeaderNames[_propertyName] = columnName;
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> Ignore()
        {
            if (_options.UnmappedProperties == null)
            {
                _options.UnmappedProperties = new List<string>();
            }

            _options.UnmappedProperties.Add(_propertyName);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> Optional(bool isOptional)
        {
            if (isOptional)
            {
                _options.OptionalProperties.Add(_propertyName);
                _options.ExplicitlyRequiredProperties.Remove(_propertyName);
            }
            else
            {
                _options.OptionalProperties.Remove(_propertyName);
                _options.ExplicitlyRequiredProperties.Add(_propertyName);
            }
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsesCustomValidator(Func<object, bool> validator, string message)
        {
            var excelCellValidator = new ExcelCellValidator
            {
                Message = message,
                Validator = validator
            };
            _options.Validations[_propertyName].Add(excelCellValidator);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapsToRowNumber()
        {
            _options.RowNumberProperty = _propertyName;
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsesColumnNumber(int i)
        {
            ExcelPropertyConfiguration.UsesColumnNumber(i, _propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsesColumnLetter(string columnLetter)
        {
            _options.CustomHeaderNumbers[_propertyName] = CellRef.ColumnNameToNumber(columnLetter) - 1;
            return _optionsBuilder;
        }

        private void CreateEntryIfNotExists()
        {
            if (!_options.Validations.ContainsKey(_propertyName))
            {
                _options.Validations[_propertyName] = new List<ExcelCellValidator>();
            }
        }

        private void CreateDictionaryIfNotExists()
        {
            if (_options.Validations == null)
            {
                _options.Validations = new Dictionary<string, List<ExcelCellValidator>>();
            }
        }
    }
}