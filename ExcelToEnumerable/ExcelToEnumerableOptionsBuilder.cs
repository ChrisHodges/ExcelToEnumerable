using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization.Json;
using ExcelToEnumerable.Attributes;
using SpreadsheetCellRef;

namespace ExcelToEnumerable
{
    internal class ExcelToEnumerableOptionsBuilder<T> : IExcelToEnumerableOptionsBuilder<T>
    {
        private readonly ExcelToEnumerableOptions<T> _options = new ExcelToEnumerableOptions<T>
        {
            AllPropertiesOptionalByDefault = true,
            IgnoreColumnsWithoutMatchingProperties = true
        };

        public IExcelToEnumerableOptionsBuilder<T> StartingFromRow(int startRow)
        {
            _options.StartRow = startRow;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> OnReadingHeaderRow(Action<IDictionary<int, string>> action)
        {
            _options.OnReadingHeaderRowAction = action;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsingHeaderNames(bool usingHeaderNames)
        {
            if (!usingHeaderNames)
            {
                _options.HeaderRow = 0;
            }

            _options.UseHeaderNames = usingHeaderNames;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> BlankRowBehaviour(BlankRowBehaviour blankRowBehaviour)
        {
            _options.BlankRowBehaviour = blankRowBehaviour;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> AggregateExceptions()
        {
            _options.ExceptionHandlingBehaviour = ExceptionHandlingBehaviour.AggregateExceptions;
            _options.ExceptionList = new List<Exception>();
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> OutputExceptionsTo(IList<Exception> list)
        {
            _options.ExceptionHandlingBehaviour = ExceptionHandlingBehaviour.LogExceptions;
            _options.ExceptionList = list;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsingSheet(int i)
        {
            _options.WorksheetName = null;
            _options.WorksheetNumber = i;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> UsingSheet(string sheetName)
        {
            _options.WorksheetName = sheetName;
            _options.WorksheetNumber = null;
            return this;
        }

        public IExcelPropertyConfiguration<T, TProperty> Property<TProperty>(
            Expression<Func<T, TProperty>> propertyExpression)
        {
            var excelPropertyConfiguration =
                new ExcelPropertyConfiguration<T, TProperty>(this, GetExpressionName(propertyExpression), _options);
            return excelPropertyConfiguration;
        }

        public IExcelToEnumerableOptions<T> Build()
        {
            AddOptionsFromAttributes(this, typeof(T));
            if (_options.AllPropertiesOptionalByDefault)
            {
                foreach (var propertyInfo in typeof(T).GetProperties())
                {
                    var propertyName = propertyInfo.Name;
                    if (!_options.OptionalColumns.Contains(propertyName))
                    {
                        _options.OptionalColumns.Add(propertyName);
                    }
                }

                foreach (var explicitlyRequiredField in _options.ExplicitlyRequiredColumns)
                {
                    if (_options.OptionalColumns.Contains(explicitlyRequiredField))
                    {
                        _options.OptionalColumns.Remove(explicitlyRequiredField);
                    }
                }
            }

            return _options;
        }

        private static void MapClassLevelAttributes(ExcelToEnumerableOptionsBuilder<T> builder,
            Type type)
        {
            var attributes = type.CustomAttributes;
            foreach (var attribute in attributes)
            {
                switch (attribute.AttributeType.Name)
                {
                    case nameof(AllPropertiesMustBeMappedToColumnsAttribute):
                        builder.AllPropertiesMustBeMappedToColumns((bool) attribute.ConstructorArguments[0].Value);
                        break;
                    case nameof(AllColumnsMustBeMappedToPropertiesAttribute):
                        builder.AllColumnsMustBeMappedToProperties((bool) attribute.ConstructorArguments[0].Value);
                        break;
                    case nameof(UsingHeaderNamesAttribute):
                        builder.UsingHeaderNames((bool) attribute.ConstructorArguments[0].Value);
                        break;
                    case nameof(StartingFromRowAttribute):
                        var rowNumber = (int) attribute.ConstructorArguments[0].Value;
                        builder.StartingFromRow(rowNumber);
                        break;
                    case nameof(UsingSheetAttribute):
                        var sheetArgument = attribute.ConstructorArguments[0].Value;
                        if (sheetArgument is int argument)
                        {
                            builder.UsingSheet(argument);
                        }
                        else
                        {
                            builder.UsingSheet(sheetArgument.ToString());
                        }

                        break;
                    case nameof(HeaderOnRowAttribute):
                        var headerOnRowNumber = (int) attribute.ConstructorArguments[0].Value;
                        builder.HeaderOnRow(headerOnRowNumber);
                        break;
                    case nameof(EndingWithRowAttribute):
                        builder.EndingWithRow((int) attribute.ConstructorArguments[0].Value);
                        break;
                    case nameof(AggregateExceptionsAttribute):
                        builder.AggregateExceptions();
                        break;
                    case nameof(BlankRowBehaviourAttribute):
                        builder.BlankRowBehaviour((BlankRowBehaviour) attribute.ConstructorArguments[0].Value);
                        break;
                    case nameof(RelaxedNumberMatchingAttribute):
                        builder.RelaxedNumberMatching((bool)attribute.ConstructorArguments[0].Value);
                        break;
                }
            }
        }

        internal static void AddOptionsFromAttributes(ExcelToEnumerableOptionsBuilder<T> builder, Type type)
        {
            MapClassLevelAttributes(builder, type);

            var properties = type.GetProperties().Distinct().ToArray();
            foreach (var property in properties)
            {
                var propertyAttributes = property.CustomAttributes.ToArray();
                foreach (var propertyAttribute in propertyAttributes)
                {
                    switch (propertyAttribute.AttributeType.Name)
                    {
                        case nameof(MapsToColumnNumberAttribute):
                            var columnNumber = (int) propertyAttribute.ConstructorArguments[0].Value;
                            ExcelPropertyConfiguration.MapsToColumnNumber(columnNumber, property.Name, builder._options);
                            break;
                        case nameof(MapsToColumnLetterAttribute):
                            var columnLetter = (string) propertyAttribute.ConstructorArguments[0].Value;
                            ExcelPropertyConfiguration.MapsToColumnNumber(CellRef.ColumnNameToNumber(columnLetter),
                                property.Name, builder._options);
                            break;
                        case nameof(OptionalColumnAttribute):
                            ExcelPropertyConfiguration.OptionalColumn(true, property.Name, builder._options);
                            break;
                        case nameof(MapFromColumnsAttribute):
                            var columnNames =
                                ((ReadOnlyCollection<CustomAttributeTypedArgument>) propertyAttribute
                                    .ConstructorArguments[0].Value).Select(x => x.Value.ToString());
                            ExcelPropertyConfiguration.MapFromColumns(columnNames, property.Name, builder._options);
                            break;
                        case nameof(MapsToColumnNamedAttribute):
                            var columnName = (string) propertyAttribute.ConstructorArguments[0].Value;
                            ExcelPropertyConfiguration.MapsToColumnNamed(columnName, property.Name, builder._options);
                            break;
                        case nameof(IgnoreColumnAttribute):
                            ExcelPropertyConfiguration.Ignore(property.Name, builder._options);
                            break;
                        case nameof(MapsToRowNumberAttribute):
                            ExcelPropertyConfiguration.MapsToRowNumber(property.Name, builder._options);
                            break;
                        case nameof(ShouldBeLessThanAttribute):
                            var maxValue = (double) propertyAttribute.ConstructorArguments[0].Value;
                            ExcelPropertyConfiguration.ShouldBeLessThan(maxValue, property.Name, builder._options);
                            break;
                        case nameof(ShouldBeGreaterThanAttribute):
                            var minValue = (double) propertyAttribute.ConstructorArguments[0].Value;
                            ExcelPropertyConfiguration.ShouldBeGreaterThan(minValue, property.Name, builder._options);
                            break;
                        case nameof(NotNullAttribute):
                            ExcelPropertyConfiguration.NotNullProperties(property.Name, builder._options);
                            break;
                        case nameof(ShouldBeOneOfAttribute):
                            // CSH 27112020 We're adding a validator directly here, rather than going via the static ExcelPropertyConfiguration because the validator we're using here using
                            // an enumerable of type object rather than type TProperty (since enforcing argument types at compile time is not possible with Attributes)
                            var objectList =
                                ((ReadOnlyCollection<CustomAttributeTypedArgument>) propertyAttribute
                                    .ConstructorArguments[0].Value).Select(x => x.Value);
                            builder._options.Validations[property.Name]
                                .Add(ExcelCellValidatorFactory.CreateShouldBeOneOf(objectList));
                            break;
                        case nameof(UniqueAttribute):
                            ExcelPropertyConfiguration.Unique(property.Name, builder._options);
                            break;
                        case nameof(RequiredColumnAttribute):
                            ExcelPropertyConfiguration.RequiredColumn(true, property.Name, builder._options);
                            break;
                    }
                }
            }
        }

        public IExcelToEnumerableOptionsBuilder<T> AllColumnsMustBeMappedToProperties(bool b)
        {
            _options.IgnoreColumnsWithoutMatchingProperties = !b;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> AllPropertiesMustBeMappedToColumns(bool b)
        {
            _options.AllPropertiesOptionalByDefault = !b;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> RelaxedNumberMatching(bool b)
        {
            _options.RelaxedNumberMatching = b;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> EndingWithRow(int maxRow)
        {
            _options.EndRow = maxRow;
            return this;
        }

        public IExcelToEnumerableOptionsBuilder<T> HeaderOnRow(int rowNumber)
        {
            _options.HeaderRow = rowNumber;
            return this;
        }

        private static string GetExpressionName<TRequiredField>(
            Expression<Func<T, TRequiredField>> isRequiredExpression)
        {
            return ((MemberExpression) isRequiredExpression.Body).Member.Name;
        }
    }
}