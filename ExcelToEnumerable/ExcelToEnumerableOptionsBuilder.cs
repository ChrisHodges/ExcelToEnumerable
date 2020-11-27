using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
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
                    if (!_options.OptionalProperties.Contains(propertyName))
                    {
                        _options.OptionalProperties.Add(propertyName);
                    }
                }

                foreach (var explicitlyRequiredField in _options.ExplicitlyRequiredProperties)
                {
                    if (_options.OptionalProperties.Contains(explicitlyRequiredField))
                    {
                        _options.OptionalProperties.Remove(explicitlyRequiredField);
                    }
                }
            }
            
            return _options;
        }

        internal static void AddOptionsFromAttributes(ExcelToEnumerableOptionsBuilder<T> builder, Type type)
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
                        var rowNumber = (int)attribute.ConstructorArguments[0].Value;
                        builder.StartingFromRow(rowNumber);
                        break;
                    case nameof(UsingSheetAttribute):
                        var sheetName = (string) attribute.ConstructorArguments[0].Value;
                        builder.UsingSheet(sheetName);
                        break;
                    case nameof(HeaderOnRowAttribute):
                        var headerOnRowNumber = (int)attribute.ConstructorArguments[0].Value;
                        builder.HeaderOnRow(headerOnRowNumber);
                        break;
                }

                foreach (var property in type.GetProperties())
                {
                    var propertyAttributes = property.CustomAttributes;
                    foreach (var propertyAttribute in propertyAttributes)
                    {
                        switch (propertyAttribute.AttributeType.Name)
                        {
                            case nameof(UsesColumnNumberAttribute):
                                var columnNumber = (int)propertyAttribute.ConstructorArguments[0].Value;
                                ExcelPropertyConfiguration.UsesColumnNumber(columnNumber, property.Name, builder._options);
                                break;
                            case nameof(UsesColumnLetterAttribute):
                                var columnLetter = (string)propertyAttribute.ConstructorArguments[0].Value;
                                ExcelPropertyConfiguration.UsesColumnNumber(CellRef.ColumnNameToNumber(columnLetter), property.Name, builder._options);
                                break;
                            case nameof(OptionalAttribute):
                                ExcelPropertyConfiguration.Optional(true, property.Name, builder._options);
                                break;
                            case nameof(MapFromColumnsAttribute):
                                var columnNames = ((ReadOnlyCollection<CustomAttributeTypedArgument>)propertyAttribute.ConstructorArguments[0].Value).Select(x => x.Value.ToString());
                                ExcelPropertyConfiguration.MapFromColumns(columnNames, property.Name, builder._options);
                                break;
                        }
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