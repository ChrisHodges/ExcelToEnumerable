using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToEnumerable
{
    internal class ExcelPropertyConfiguration<T, TProperty> : IExcelPropertyConfiguration<T, TProperty>
    {
        private readonly IExcelToEnumerableOptions<T> _options;
        private readonly string _propertyName;
        private IExcelToEnumerableOptionsBuilder<T> _optionsBuilder;

        public ExcelPropertyConfiguration(IExcelToEnumerableOptionsBuilder<T> optionsBuilder, string propertyName, IExcelToEnumerableOptions<T> options)
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

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(params TProperty[] oneOf)
        {
            _options.Validations[_propertyName].Add(ExcelCellValidatorFactory.CreateShouldBeOneOf(oneOf));
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapFromColumns(IEnumerable<string> columnNames)
        {
            if (_options.CollectionConfigurations == null)
            {
                _options.CollectionConfigurations = new Dictionary<string, ExcelToEnumerableCollectionConfiguration>();
            }

            var configuration = new ExcelToEnumerableCollectionConfiguration
            {
                PropertyName = _propertyName,
                ColumnNames = columnNames
            };

            _options.CollectionConfigurations.Add(_propertyName, configuration);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapFromColumns(params string[] columnNames)
        {
            return MapFromColumns(columnNames.ToList());
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

        public IExcelToEnumerableOptionsBuilder<T> ShouldBeUnique()
        {
            if (_options.UniqueFields == null)
            {
                _options.UniqueFields = new List<string>();
            }
            _options.UniqueFields.Add(_propertyName);
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
            if (_options.SkippedFields == null)
            {
                _options.SkippedFields = new List<string>();
            }
            _options.SkippedFields.Add(_propertyName);
            return _optionsBuilder;
        }

        public IExcelToEnumerableOptionsBuilder<T> MapsToRowNumber()
        {
            _options.RowNumberColumn = _propertyName;
            return _optionsBuilder;
        }
    }
}