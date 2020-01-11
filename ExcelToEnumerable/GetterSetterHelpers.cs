using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelToEnumerable
{
    public static class GetterSetterHelpers
    {
        public static Func<object, object> GetGetter(PropertyInfo propertyInfo)
        {
            var method = propertyInfo.GetMethod;
            var obj = Expression.Parameter(typeof(object), "obj");

            var expr =
                Expression.Lambda<Func<object, object>>(
                    Expression.Convert(
                        Expression.Call(
                            Expression.Convert(obj, method.DeclaringType),
                            method),
                        typeof(object)),
                    obj);

            return expr.Compile();
        }

        public static Action<object, object> GetSetter(PropertyInfo propertyInfo)
        {
            var instance = Expression.Parameter(typeof(object), "instance");
            var argument = Expression.Parameter(typeof(object), "value");
            var setterCall = Expression.Call(
                Expression.Convert(instance, propertyInfo.DeclaringType),
                propertyInfo.GetSetMethod(true), 
                Expression.Convert(argument, propertyInfo.PropertyType));
            var setter = (Action<object, object>) Expression.Lambda(setterCall, instance, argument).Compile();
            return setter;
        }
        
        public static Action<object, object> GetAdder(PropertyInfo propertyInfo)
        {
            var getter = GetGetter(propertyInfo);
            var setter = GetSetter(propertyInfo);
            var collectionCreator = GetCollectionCreator(propertyInfo);
            var collectionAdder = GetCollectionAdder(propertyInfo);
            Action<object, object> expr = (o, o1) =>
            {
                var list = getter(o);
                if (list == null)
                {
                    list = collectionCreator();
                    setter(o, list);
                }
                collectionAdder(list, o1);
            };
            return expr;
        }


        internal static Func<object> GetCollectionCreator(PropertyInfo propertyInfo)
        {
            var constructorInfo = propertyInfo.PropertyType.GetConstructor(new Type[0]);
            var newExpression = Expression.New(constructorInfo);
            var expr = Expression.Lambda<Func<object>>(Expression.Convert(newExpression, typeof(object))).Compile();
            return expr;
        }

        private static Action<object, object> GetCollectionAdder(PropertyInfo propertyInfo)
        {
            var instance = Expression.Parameter(typeof(object), "instance");
            var argument = Expression.Parameter(typeof(object), "value");
            var addMethodInfo = propertyInfo.PropertyType.GetMethod("Add");
            var collectionGenericType = propertyInfo.PropertyType.GetGenericArguments().First();
            var methodCall = Expression.Call(Expression.Convert(instance, propertyInfo.PropertyType), addMethodInfo, Expression.Convert(argument, collectionGenericType));
            var expr = Expression.Lambda<Action<object, object>>(methodCall, instance, argument).Compile();
            return expr;
        }

        private static Action<object, object, object> GetDictionaryItemAdder(PropertyInfo propertyInfo)
        {
            var instance = Expression.Parameter(typeof(object), "instance");
            var key = Expression.Parameter(typeof(object), "key");
            var value = Expression.Parameter(typeof(object), "value");
            var addMethodInfo = propertyInfo.PropertyType.GetMethod("Add");
            var genericArguments = propertyInfo.PropertyType.GetGenericArguments();
            var keyGenericType = genericArguments.First();
            var valueGenericType = genericArguments.Skip(1).First();
            var methodCall = Expression.Call(Expression.Convert(instance, propertyInfo.PropertyType), addMethodInfo, Expression.Convert(key, keyGenericType), Expression.Convert(value, valueGenericType));
            var expr = Expression.Lambda<Action<object, object, object>>(methodCall, instance, key, value).Compile();
            return expr;
        }

        public static Action<object, object, object> GetDictionaryAdder(PropertyInfo propertyInfo)
        {
            var getter = GetGetter(propertyInfo);
            var setter = GetSetter(propertyInfo);
            var collectionCreator = GetCollectionCreator(propertyInfo);
            var collectionAdder = GetDictionaryItemAdder(propertyInfo);
            Action<object, object, object> expr = (o, oKey, oValue) =>
            {
                var list = getter(o);
                if (list == null)
                {
                    list = collectionCreator();
                    setter(o, list);
                }
                collectionAdder(list, oKey, oValue);
            };
            return expr;
        }

        public static Action<object, object> GetDictionaryAdder(PropertyInfo propertyInfo, object key)
        {
            var getter = GetGetter(propertyInfo);
            var setter = GetSetter(propertyInfo);
            var collectionCreator = GetCollectionCreator(propertyInfo);
            var collectionAdder = GetDictionaryItemAdder(propertyInfo);
            Action<object, object> expr = (o, oValue) =>
            {
                var list = getter(o);
                if (list == null)
                {
                    list = collectionCreator();
                    setter(o, list);
                }
                collectionAdder(list, key, oValue);
            };
            return expr;
        }
    }
}