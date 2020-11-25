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

        private readonly ConcurrentDictionary<int, FromRowConstructor>
            _fromRowConstructorDictionary =
                new ConcurrentDictionary<int, FromRowConstructor>();

        private ExcelToEnumerableContext()
        {
        }

        public static ExcelToEnumerableContext Instance => Lazy.Value;

        public FromRowConstructor GetFromRowConstructor<T>(IExcelToEnumerableOptions<T> options)
        {
            var hashCode = options.GetHashCode();
            return _fromRowConstructorDictionary.ContainsKey(hashCode)
                ? _fromRowConstructorDictionary[hashCode]
                : null;
        }

        public FromRowConstructor SetFromRowConstructor<T>(IExcelToEnumerableOptions<T> options)
        {
            var hashCode = options.GetHashCode();
            if (_fromRowConstructorDictionary.ContainsKey(hashCode))
            {
                return _fromRowConstructorDictionary[hashCode];
            }

            var fromRowConstructor = CreateFromRowConstructorFromPropertyDescriptorCollection(options);
            _fromRowConstructorDictionary.TryAdd(hashCode, fromRowConstructor);
            return fromRowConstructor;
        }

        internal IEnumerable<FromCellSetter> GetSetters<T>(Type type, IExcelToEnumerableOptions<T> options)
        {
            var propertyDescriptorCollection = TypeDescriptor.GetProperties(type);
            var list = propertyDescriptorCollection.Cast<PropertyDescriptor>();
            if (options.UnmappedProperties != null)
            {
                list = list.Where(x => !options.UnmappedProperties.Contains(x.Name));
            }

            return list.Select(x =>
                    GetSettersForProperty(type.GetProperty(x.Name), Array.IndexOf(list.ToArray(), x), options))
                .SelectMany(y => y).ToArray();
        }
        
        private IEnumerable<FromCellSetter> GetSettersForProperty<T>(PropertyInfo propertyInfo, int index,
            IExcelToEnumerableOptions<T> options)
        {
            if (propertyInfo.PropertyType != typeof(string) &&
                typeof(IEnumerable).IsAssignableFrom(propertyInfo.PropertyType))
            {
                if (options.CollectionConfigurations == null ||
                    !options.CollectionConfigurations.ContainsKey(propertyInfo.Name))
                {
                    return new FromCellSetter[0];
                }

                return GetSettersForEnumerable(propertyInfo, index, options);
            }

            var setter = GetterSetterHelpers.GetSetter(propertyInfo);
            var getter = options.UniqueProperties != null && options.UniqueProperties.Contains(propertyInfo.Name)
                ? GetterSetterHelpers.GetGetter(propertyInfo)
                : null;
            var fromCellSetter = new FromCellSetter
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
                        : null
            };
            if (propertyInfo.Name == options.RowNumberProperty)
            {
                fromCellSetter.ColumnName = null;
            }

            if (options.RequiredFields.Contains(propertyInfo.Name))
            {
                fromCellSetter.Validators = new List<ExcelCellValidator> {ExcelCellValidatorFactory.CreateRequired()};
            }

            return new[] {fromCellSetter};
        }

        private IEnumerable<FromCellSetter> GetSettersForEnumerable<T>(PropertyInfo propertyInfo, int index,
            IExcelToEnumerableOptions<T> options)
        {
            var collectionsConfig = options.CollectionConfigurations[propertyInfo.Name];
            var isDictionary = typeof(IDictionary).IsAssignableFrom(propertyInfo.PropertyType);
            var enumerableType = propertyInfo.PropertyType.GenericTypeArguments[0];
            var fromCellSetters = collectionsConfig.ColumnNames.Select(x => new FromCellSetter
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

        internal FromRowConstructor CreateFromRowConstructorFromPropertyDescriptorCollection<T>(
            IExcelToEnumerableOptions<T> options)
        {
            var type = typeof(T);
            var constructor = new FromRowConstructor
            {
                Type = type,
                Setters = GetSetters(type, options),
            };
            return constructor;
        }

        public void CreateMapper<T>(IExcelToEnumerableOptions<T> options)
        {
            var fromRowConstructor = CreateFromRowConstructorFromPropertyDescriptorCollection(options);
            _fromRowConstructorDictionary.TryAdd(options.GetHashCode(), fromRowConstructor);
        }
    }
}