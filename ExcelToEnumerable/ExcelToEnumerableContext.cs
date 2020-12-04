using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ExcelToEnumerable
{
    internal sealed class ExcelToEnumerableContext : IExcelToEnumerableContext
    {
        private static readonly Lazy<ExcelToEnumerableContext>
            Lazy =
                new Lazy<ExcelToEnumerableContext>
                    (() => new ExcelToEnumerableContext());

        private readonly ConcurrentDictionary<int, RowMapper>
            _rowMapperDictionary =
                new ConcurrentDictionary<int, RowMapper>();

        private ExcelToEnumerableContext()
        {
        }

        public static ExcelToEnumerableContext Instance => Lazy.Value;

        public RowMapper GetRowMapper<T>(IExcelToEnumerableOptions<T> options)
        {
            var hashCode = options.GetHashCode();
            return _rowMapperDictionary.ContainsKey(hashCode)
                ? _rowMapperDictionary[hashCode]
                : null;
        }

        public RowMapper SetRowMapper<T>(IExcelToEnumerableOptions<T> options)
        {
            var hashCode = options.GetHashCode();
            if (_rowMapperDictionary.ContainsKey(hashCode))
            {
                return _rowMapperDictionary[hashCode];
            }

            var rowMapper = CreateRowMapperFromPropertyDescriptorCollection(options);
            _rowMapperDictionary.TryAdd(hashCode, rowMapper);
            return rowMapper;
        }

        private IEnumerable<PropertySetter> GetSetters<T>(Type type, IExcelToEnumerableOptions<T> options)
        {
            var propertyDescriptorCollection = TypeDescriptor.GetProperties(type);
            var propertyDescriptors = propertyDescriptorCollection.Cast<PropertyDescriptor>().ToArray();
            if (options.UnmappedProperties != null)
            {
                propertyDescriptors = propertyDescriptors.Where(x => !options.UnmappedProperties.Contains(x.Name)).ToArray();
            }

            return propertyDescriptors.Select(x =>
                    GetSettersForProperty(type.GetProperty(x.Name), Array.IndexOf(propertyDescriptors.ToArray(), x), options))
                .SelectMany(y => y).ToArray();
        }

        private IEnumerable<PropertySetter> GetSettersForProperty<T>(PropertyInfo propertyInfo, int index,
            IExcelToEnumerableOptions<T> options)
        {
            if (propertyInfo.PropertyType != typeof(string) &&
                typeof(IEnumerable).IsAssignableFrom(propertyInfo.PropertyType))
            {
                if (options.CollectionConfigurations == null ||
                    !options.CollectionConfigurations.ContainsKey(propertyInfo.Name))
                {
                    return new PropertySetter[0];
                }

                return GetSettersForEnumerable(propertyInfo, index, options);
            }

            var setter = GetterSetterHelpers.GetSetter(propertyInfo);
            var getter = options.UniqueProperties != null && options.UniqueProperties.Contains(propertyInfo.Name)
                ? GetterSetterHelpers.GetGetter(propertyInfo)
                : null;
            var fromCellSetter = new PropertySetter
            {
                Getter = getter,
                Setter = setter,
                Type = propertyInfo.PropertyType,
                ColumnName =
                    options.CustomHeaderNames != null && options.CustomHeaderNames.ContainsKey(propertyInfo.Name)
                        ? options.CustomHeaderNames[propertyInfo.Name]
                        : propertyInfo.Name.ToLowerInvariant(),
                PropertyName = propertyInfo.Name,
                CustomMapping = options.CustomMappings != null && options.CustomMappings.ContainsKey(propertyInfo.Name)
                    ? options.CustomMappings[propertyInfo.Name]
                    : DefaultTypeMappers.Dictionary.ContainsKey(propertyInfo.PropertyType)
                        ? DefaultTypeMappers.Dictionary[propertyInfo.PropertyType]
                        : null,
                RelaxedNumberMatching = options.RelaxedNumberMatching && propertyInfo.PropertyType.IsNumeric()
            };
            if (propertyInfo.Name == options.RowNumberProperty)
            {
                fromCellSetter.ColumnName = null;
            }

            if (options.NotNullProperties.Contains(propertyInfo.Name))
            {
                fromCellSetter.Validators = new List<ExcelCellValidator> {ExcelCellValidatorFactory.CreateRequired()};
            }

            return new[] {fromCellSetter};
        }

        private IEnumerable<PropertySetter> GetSettersForEnumerable<T>(PropertyInfo propertyInfo, int index,
            IExcelToEnumerableOptions<T> options)
        {
            var collectionsConfig = options.CollectionConfigurations[propertyInfo.Name];
            var isDictionary = typeof(IDictionary).IsAssignableFrom(propertyInfo.PropertyType);
            var enumerableType = propertyInfo.PropertyType.GenericTypeArguments[0];
            var fromCellSetters = collectionsConfig.ColumnNames.Select(x => new PropertySetter
            {
                ColumnName = x.ToLowerInvariant(),
                PropertyName = propertyInfo.Name,
                Setter = isDictionary
                    ? GetterSetterHelpers.GetDictionaryAdder(propertyInfo, x)
                    : GetterSetterHelpers.GetAdder(propertyInfo),
                Type = enumerableType,
                CustomMapping = options.CustomMappings != null && options.CustomMappings.ContainsKey(propertyInfo.Name)
                    ? options.CustomMappings[propertyInfo.Name]
                    : DefaultTypeMappers.Dictionary.ContainsKey(enumerableType)
                        ? DefaultTypeMappers.Dictionary[enumerableType]
                        : null
            });
            return fromCellSetters;
        }

        private RowMapper CreateRowMapperFromPropertyDescriptorCollection<T>(
            IExcelToEnumerableOptions<T> options)
        {
            var type = typeof(T);
            var constructor = new RowMapper
            {
                Setters = GetSetters(type, options),
            };
            return constructor;
        }

        public void CreateMapper<T>(IExcelToEnumerableOptions<T> options)
        {
            var fromRowConstructor = CreateRowMapperFromPropertyDescriptorCollection(options);
            _rowMapperDictionary.TryAdd(options.GetHashCode(), fromRowConstructor);
        }
    }
}