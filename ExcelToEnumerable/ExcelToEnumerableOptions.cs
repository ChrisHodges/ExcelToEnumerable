using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToEnumerable
{
    internal class ExcelToEnumerableOptions<T> : IExcelToEnumerableOptions<T>
    {
        private Dictionary<string, string> _customHeaderNames;
        private Dictionary<string, Func<object, object>> _customMappings;
        private IEnumerable<string> _loweredRequiredColumns;
        private List<string> _requiredFields = new List<string>();
        private string rowNumberColumn;

        public ExcelToEnumerableOptions()
        {
            WorksheetNumber = 0;
            BlankRowBehaviour = BlankRowBehaviour.Ignore;
            ExceptionHandlingBehaviour = ExceptionHandlingBehaviour.ThrowOnFirstException;
            UseHeaderNames = true;
        }

        public bool AllPropertiesOptionalByDefault { get; set; }

        public Dictionary<string, List<ExcelCellValidator>> Validations { get; set; }

        public bool UseHeaderNames { get; set; }

        public int StartRow { get; set; }

        public BlankRowBehaviour BlankRowBehaviour { get; internal set; }

        public ExceptionHandlingBehaviour ExceptionHandlingBehaviour { get; set; }
        public Type MappedType { get; set; }

        public IList<string> RequiredFields
        {
            get => _requiredFields;
            set => _requiredFields = value.ToList();
        }

        public IList<Exception> ExceptionList { get; internal set; }

        public IEnumerable<string> LoweredRequiredColumns
        {
            get
            {
                if (_loweredRequiredColumns == null)
                {
                    _loweredRequiredColumns = CreateLoweredRequiredColumns();
                }

                return _loweredRequiredColumns;
            }
        }

        public Dictionary<string, ExcelToEnumerableCollectionConfiguration> CollectionConfigurations { get; set; }

        public string WorksheetName { get; set; }

        public int? WorksheetNumber { get; set; }
        public List<string> UniqueFields { get; set; }
        public int? EndRow { get; internal set; }
        public int HeaderRow { get; internal set; }
        public Action<IDictionary<int, string>> OnReadingerHeaderRowAction { get; internal set; }

        public Dictionary<string, Func<object, object>> CustomMappings
        {
            get
            {
                if (_customMappings == null)
                {
                    _customMappings = new Dictionary<string, Func<object, object>>();
                }

                return _customMappings;
            }
        }

        public Dictionary<string, string> CustomHeaderNames
        {
            get
            {
                if (_customHeaderNames == null)
                {
                    _customHeaderNames = new Dictionary<string, string>();
                }

                return _customHeaderNames;
            }
        }

        public List<string> SkippedFields { get; set; }

        public string RowNumberColumn
        {
            get => rowNumberColumn;
            set
            {
                LoweredRowNumberColumn = value.ToLowerInvariant();
                rowNumberColumn = value;
            }
        }

        public string LoweredRowNumberColumn { get; private set; }
        public Dictionary<string, int> CustomHeaderNumbers { get; } = new Dictionary<string, int>();
        public List<string> OptionalFields { get; set; } = new List<string>();
        public bool IgnoreColumnsWithoutMatchingProperties { get; set; }
        public List<string> ExplictlyRequiredFields { get; set; } = new List<string>();

        private IEnumerable<string> CreateLoweredRequiredColumns()
        {
            var requiredFields = RequiredFields.Select(x => x.ToLowerInvariant());
            if (CollectionConfigurations == null && CustomHeaderNames == null)
            {
                return requiredFields;
            }

            var fieldList = RequiredFields.Select(x => x.ToLowerInvariant()).ToList();
            if (CollectionConfigurations != null)
            {
                foreach (var collectionConfig in CollectionConfigurations.Where(x =>
                    requiredFields.Contains(x.Value.PropertyName.ToLowerInvariant())))
                {
                    fieldList.Remove(collectionConfig.Value.PropertyName.ToLowerInvariant());
                    fieldList.AddRange(collectionConfig.Value.ColumnNames.Select(x => x.ToLowerInvariant()));
                }
            }

            if (CustomHeaderNames != null)
            {
                foreach (var item in CustomHeaderNames)
                {
                    fieldList.Remove(item.Key.ToLowerInvariant());
                    fieldList.Add(item.Value);
                }
            }

            return fieldList;
        }

        public void AddRequiredField(string requiredField)
        {
            _requiredFields.Add(requiredField);
        }
    }
}