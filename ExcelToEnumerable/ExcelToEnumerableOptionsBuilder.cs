using System;
using System.Collections.Generic;
using System.Linq.Expressions;

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
        
        public IExcelToEnumerableOptionsBuilder<T> IgnoreColumnsWithoutMatchingProperties(bool b)
        {
            _options.IgnoreColumnsWithoutMatchingProperties = b;
            return this;
        }
        
        public IExcelToEnumerableOptionsBuilder<T> IgnorePropertiesWithoutMatchingColumns(bool b)
        {
            _options.AllPropertiesOptionalByDefault = b;
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