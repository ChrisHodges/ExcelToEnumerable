using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class GetterSetterHelperTests
    {
        [Fact]
        public void GetGetterWorks()
        {
            var propertyInfo = typeof(CollectionTestClass).GetProperty("Collection");
            var getter = GetterSetterHelpers.GetGetter( propertyInfo);
            var testClass = new CollectionTestClass();
            var result = getter(testClass);
            result.Should().BeNull();
            testClass.Collection = new List<string>();
            result = getter(testClass);
            result.Should().NotBeNull();
        }

        [Fact]
        public void GetAdderWorks()
        {
            var propertyInfo = typeof(CollectionTestClass).GetProperty("Collection");
            var adder = GetterSetterHelpers.GetAdder(propertyInfo);
            var testClass = new CollectionTestClass();
            adder(testClass, "A");
            adder(testClass, "B");
            testClass.Collection.First().Should().Be("A");
            testClass.Collection.Last().Should().Be("B");
        }

        [Fact]
        public void GetIntAdderWorks()
        {
            var propertyInfo = typeof(CollectionTestClass).GetProperty("IntCollection");
            var adder = GetterSetterHelpers.GetAdder(propertyInfo);
            var testClass = new CollectionTestClass();
            adder(testClass, 1);
            adder(testClass, 2);
            testClass.IntCollection.First().Should().Be(1);
            testClass.IntCollection.Last().Should().Be(2);
        }

        [Fact]
        public void GetCollectionCreatorWorks()
        {
            var propertyInfo = typeof(CollectionTestClass).GetProperty("Collection");
            var collectionCreator = GetterSetterHelpers.GetCollectionCreator(propertyInfo);
            var result = (List<string>) collectionCreator();
            result.Should().NotBeNull();
        }

        [Fact]
        public void GetDictionaryAdderWorks()
        {
            var propertyInfo = typeof(DictionaryCollectionTestClass).GetProperty("Collection");
            var testClass = new DictionaryCollectionTestClass();
            var adder = GetterSetterHelpers.GetDictionaryAdder(propertyInfo);
            adder(testClass, "a", "b");
            testClass.Collection["a"].Should().Be("b");
        }

        [Fact]
        public void GetKeySpecifiedDictionaryAdderWorks()
        {
            var propertyInfo = typeof(DictionaryCollectionTestClass).GetProperty("Collection");
            var testClass = new DictionaryCollectionTestClass();
            var adder = GetterSetterHelpers.GetDictionaryAdder(propertyInfo, "z");
            adder(testClass, "y");
            testClass.Collection["z"].Should().Be("y");
        }
    }
}