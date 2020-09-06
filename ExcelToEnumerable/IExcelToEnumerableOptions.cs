using System;
using System.Collections.Generic;

namespace ExcelToEnumerable
{
    public interface IExcelToEnumerableOptions<T> : IExcelToEnumerableOptions
    {
        BlankRowBehaviour BlankRowBehaviour { get; }
        IList<Exception> ExceptionList { get; }
        Type MappedType { get; set; }
        IList<string> RequiredFields { get; }
        int? WorksheetNumber { get; }
        string WorksheetName { get; }
        bool UseHeaderNames { get; }
        int StartRow { get; }
        int HeaderRow { get; }
        int? EndRow { get; }

        ExceptionHandlingBehaviour ExceptionHandlingBehaviour { get; }
        Dictionary<string, List<ExcelCellValidator>> Validations { get; set; }
        IEnumerable<string> LoweredRequiredColumns { get; }
        Dictionary<string, ExcelToEnumerableCollectionConfiguration> CollectionConfigurations { get; set; }
        List<string> UniqueFields { get; set; }
        Action<IDictionary<int, string>> OnReadingerHeaderRowAction { get; }
        Dictionary<string, Func<object, object>> CustomMappings { get; }
        Dictionary<string,string> CustomHeaderNames { get; }
        List<string> SkippedFields { get; set; }
        string RowNumberColumn { get; set; }
        string LoweredRowNumberColumn { get; }
        Dictionary<string,int> CustomHeaderNumbers { get;}
        List<string> OptionalFields { get; set; }
        bool IgnoreUnmappedColumns { get; set; }
    }

    public interface IExcelToEnumerableOptions
    {
    }
}