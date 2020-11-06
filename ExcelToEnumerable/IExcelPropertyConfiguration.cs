using System;
using System.Collections.Generic;

namespace ExcelToEnumerable
{
    public interface IExcelPropertyConfiguration<T, TProperty>
    {
        IExcelToEnumerableOptionsBuilder<T> ShouldBeGreaterThan(double minValue);
        IExcelToEnumerableOptionsBuilder<T> ShouldBeLessThan(double maxValue);
        IExcelToEnumerableOptionsBuilder<T> ShouldBeUnique();
        IExcelToEnumerableOptionsBuilder<T> IsRequired();

        IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(IEnumerable<TProperty> minValue);

        IExcelToEnumerableOptionsBuilder<T> ShouldBeOneOf(params TProperty[] permissiableValues);
        IExcelToEnumerableOptionsBuilder<T> MapFromColumns(IEnumerable<string> columnNames);
        IExcelToEnumerableOptionsBuilder<T> MapFromColumns(params string[] columnNames);
        IExcelToEnumerableOptionsBuilder<T> UsesCustomMapping(Func<object, object> mapping);
        IExcelToEnumerableOptionsBuilder<T> UsesColumnNamed(string columnName);
        IExcelToEnumerableOptionsBuilder<T> Ignore();
        IExcelToEnumerableOptionsBuilder<T> MapsToRowNumber();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="i">The ZERO-INDEXED column number</param>
        /// <returns></returns>
        IExcelToEnumerableOptionsBuilder<T> UsesColumnNumber(int i);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnLetter">
        /// e.g. A, M, AB, ZZA, etc.</param>
        IExcelToEnumerableOptionsBuilder<T> UsesColumnLetter(string columnLetter);

        IExcelToEnumerableOptionsBuilder<T> Optional(bool isOptional = true);
        IExcelToEnumerableOptionsBuilder<T> UsesCustomValidator(ExcelCellValidator excelCellValidator);
    }
}