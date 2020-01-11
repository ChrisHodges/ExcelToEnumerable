using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelToEnumerable
{
    public interface IExcelToEnumerableOptionsBuilder<T>
    {
        IExcelToEnumerableOptionsBuilder<T> StartingFromRow(int startRow);
        IExcelToEnumerableOptionsBuilder<T> UsingHeaderNames(bool usingHeaderNames);
        IExcelToEnumerableOptionsBuilder<T> BlankRowBehaviour(BlankRowBehaviour blankRowBehaviour);
        IExcelToEnumerableOptionsBuilder<T> AggregateExceptions();
        IExcelToEnumerableOptionsBuilder<T> OutputExceptionsTo(IList<Exception> list);

        IExcelToEnumerableOptionsBuilder<T> UsingSheet(int i);
        IExcelToEnumerableOptionsBuilder<T> UsingSheet(string sheetname);

        IExcelPropertyConfiguration<T, TProperty>
            Property<TProperty>(Expression<Func<T, TProperty>> propertyExpression);
        IExcelToEnumerableOptionsBuilder<T> EndingWithRow(int v);
        IExcelToEnumerableOptionsBuilder<T> HeaderOnRow(int rowNumber);
        IExcelToEnumerableOptionsBuilder<T> OnReadingHeaderRow(Action<IDictionary<int, string>> action);
        IExcelToEnumerableOptions<T> Build();
    }
}