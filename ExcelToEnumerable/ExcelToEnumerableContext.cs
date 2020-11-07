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

        private readonly ConcurrentDictionary<IExcelToEnumerableOptions, FromRowConstructor>
            _fromRowConstructorDictionary =
                new ConcurrentDictionary<IExcelToEnumerableOptions, FromRowConstructor>();

        private ExcelToEnumerableContext()
        {
        }

        public static ExcelToEnumerableContext Instance => Lazy.Value;

        public FromRowConstructor GetFromRowConstructor<T>(IExcelToEnumerableOptions<T> options)
        {
            return _fromRowConstructorDictionary.ContainsKey(options)
                ? _fromRowConstructorDictionary[options]
                : null;
        }

        public FromRowConstructor SetFromRowConstructor<T>(IExcelToEnumerableOptions<T> options)
        {
            if (_fromRowConstructorDictionary.ContainsKey(options))
            {
                return _fromRowConstructorDictionary[options];
            }

            var fromRowConstructor = CreateFromRowConstructorFromPropertyDescriptorCollection(options);
            _fromRowConstructorDictionary.TryAdd(options, fromRowConstructor);
            return fromRowConstructor;
        }

        internal IEnumerable<FromCellSetter> GetSetters<T>(Type type, IExcelToEnumerableOptions<T> options)
        {
            var propertyDescriptorCollection = TypeDescriptor.GetProperties(type);
            var list = propertyDescriptorCollection.Cast<PropertyDescriptor>();
            if (options.SkippedFields != null)
            {
                list = list.Where(x => !options.SkippedFields.Contains(x.Name));
            }

            return list.Select(x =>
                    GetSettersForProperty(type.GetProperty(x.Name), Array.IndexOf(list.ToArray(), x), options))
                .SelectMany(y => y).ToArray();
        }

        /// <summary>
        ///     This usually returns a single setter apart from where the property is of type IEnumerable
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <param name="index"></param>
        /// <param name="options"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        internal IEnumerable<FromCellSetter> GetSettersForProperty<T>(PropertyInfo propertyInfo, int index,
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
            var getter = options.UniqueFields != null && options.UniqueFields.Contains(propertyInfo.Name)
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
            if (propertyInfo.Name == options.RowNumberColumn)
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
            _fromRowConstructorDictionary.TryAdd(options, fromRowConstructor);
        }
    }
}