namespace ExcelToEnumerable
{
    internal interface IExcelToEnumerableContext
    {
        RowMapper SetRowMapper<T>(IExcelToEnumerableOptions<T> options);
        RowMapper GetRowMapper<T>(IExcelToEnumerableOptions<T> options);
    }
}