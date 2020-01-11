namespace ExcelToEnumerable
{
    internal interface IExcelToEnumerableContext
    {
        FromRowConstructor SetFromRowConstructor<T>(IExcelToEnumerableOptions<T> options);
        FromRowConstructor GetFromRowConstructor<T>(IExcelToEnumerableOptions<T> options);
    }
}