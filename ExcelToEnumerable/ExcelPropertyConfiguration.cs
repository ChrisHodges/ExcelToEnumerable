using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToEnumerable.Exceptions;
using SpreadsheetCellRef;

namespace ExcelToEnumerable
{
    internal static class ExcelPropertyConfiguration
    {
        internal static void MapsToColumnNumber<T>(int i, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (i < 1)
            {
                throw new ExcelToEnumerableConfigException(
                    $"Unable to map '{propertyName}' to column {i}. MapsToColumnNumber expects a 1-based column number");
            }

            options.CustomHeaderNumbers[propertyName] = i - 1;
        }
        
        internal static void OptionalColumn<T>(bool isOptional, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (isOptional)
            {
                options.OptionalColumns.Add(propertyName);
                options.ExplicitlyRequiredColumns.Remove(propertyName);
            }
            else
            {
                options.OptionalColumns.Remove(propertyName);
                options.ExplicitlyRequiredColumns.Add(propertyName);
            }
        }
        
        internal static void RequiredColumn<T>(bool isRequired, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (isRequired)
            {
                options.ExplicitlyRequiredColumns.Add(propertyName);
            }
            else
            {
                options.ExplicitlyRequiredColumns.Remove(propertyName);
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

        public static void MapsToColumnNamed<T>(string columnName, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            options.CustomHeaderNames[propertyName] = columnName;
        }

        public static void Ignore<T>(string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (options.UnmappedProperties == null)
            {
                options.UnmappedProperties = new List<string>();
            }

            options.UnmappedProperties.Add(propertyName);
        }

        public static void MapsToRowNumber<T>(string propertyName, IExcelToEnumerableOptions<T> options)
        {
            options.RowNumberProperty = propertyName;
        }

        public static void ShouldBeLessThan<T>(double maxValue, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (!options.Validations.ContainsKey(propertyName))
            {
                options.Validations[propertyName] = new List<ExcelCellValidator>();
            }
            options.Validations[propertyName].Add(ExcelCellValidatorFactory.CreateLessThan(maxValue));
        }

        public static void ShouldBeGreaterThan<T>(double minValue, string propertyName, IExcelToEnumerableOptions<T> options)
        {
            if (!options.Validations.ContainsKey(propertyName))
            {
                options.Validations[propertyName] = new List<ExcelCellValidator>();
            }
            options.Validations[propertyName].Add(ExcelCellValidatorFactory.CreateGreaterThan(minValue));
        }

        public static void NotNullProperties<T>(string propertyName, IExcelToEnumerableOptions<T> options)
        {
            options.NotNullProperties.Add(propertyName);
        }

        public static void Unique<T>(string propertyName, IExcelToEnumerableOptions<T> options)
        {
            options.UniqueProperties.Add(propertyName);
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
            CreateEntryIfNotExists();
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeGreaterThan(double minValue)
        {
            ExcelPropertyConfiguration.ShouldBeGreaterThan(minValue, _propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeLessThan(double maxValue)
        {
            ExcelPropertyConfiguration.ShouldBeLessThan(maxValue, _propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(IEnumerable<TProperty> oneOf)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateShouldBeOneOf(oneOf));
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(params TProperty[] permissableValues)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateShouldBeOneOf(permissableValues));
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
            ExcelPropertyConfiguration.Unique(_propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> NotNull()
        {
            ExcelPropertyConfiguration.NotNullProperties(_propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsesCustomMapping(Func<object, object> mapping)
        {
            _options.CustomMappings[_propertyName] = mapping;
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapsToColumnNamed(string columnName)
        {
            ExcelPropertyConfiguration.MapsToColumnNamed(columnName, _propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> IgnoreColumn()
        {
            ExcelPropertyConfiguration.Ignore(_propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> OptionalColumn(bool isOptional)
        {
            ExcelPropertyConfiguration.OptionalColumn(isOptional, _propertyName, _options);
            return _optionsBuilder;
        }
        
        
        public IExcelToEnumerableOptionsBuilder<T> RequiredColumn(bool isRequired)
        {
            ExcelPropertyConfiguration.RequiredColumn(isRequired, _propertyName, _options);
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
            ExcelPropertyConfiguration.MapsToRowNumber(_propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapsToColumnNumber(int i)
        {
            ExcelPropertyConfiguration.MapsToColumnNumber(i, _propertyName, _options);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapsToColumnLetter(string columnLetter)
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
    }
}