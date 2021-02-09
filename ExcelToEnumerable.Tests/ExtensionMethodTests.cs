using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelToEnumerable.Exceptions;
using ExcelToEnumerable.Tests.TestClasses;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class ExtensionMethodTests
    {
        [Fact]
        public void AggregateExceptionHandlingWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet4Errors.xlsx");
            Action action = () =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                    x => x.AggregateExceptions()
                        .UsingHeaderNames(false)
                );
            };
            action.Should().Throw<AggregateException>().And.InnerExceptions.Count.Should().Be(3);
        }

        [Fact]
        public void AllPropertiesMustBeMappedToColumnsTrueWorks()
        {
            Action action = () =>
            {
                var testSpreadsheetLocation = TestHelper.TestsheetPath("IgnorePropertyNames.xlsx");
                testSpreadsheetLocation.ExcelToEnumerable<AllPropertiesMustBeMappedToColumnsTestClass>(x => x
                    .AllPropertiesMustBeMappedToColumns(true)
                );
            };
            action.Should().ThrowExactly<ExcelToEnumerableInvalidHeaderException>().WithMessage("Missing headers: 'notonspreadsheet'. ");
        }
        
        [Fact]
        public void AllPropertiesMustBeMappedToColumns_Attribute_Works()
        {
            Action action = () =>
            {
                var testSpreadsheetLocation = TestHelper.TestsheetPath("IgnorePropertyNames.xlsx");
                testSpreadsheetLocation.ExcelToEnumerable<AllPropertiesMustBeMappedToColumnsAttributeTestClass>();
            };
            action.Should().Throw<ExcelToEnumerableInvalidHeaderException>()
                .WithMessage("Missing headers: 'notonspreadsheet'. ");
        }

        [Fact]
        public void AllPropertiesMustBeMappedToColumns_False_Works()
        {

            var testSpreadsheetLocation = TestHelper.TestsheetPath("IgnorePropertyNames.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<AllPropertiesMustBeMappedToColumnsTestClass>(x => x
                .AllPropertiesMustBeMappedToColumns(false)
            ).ToArray();
            results.Length.Should().Be(1);
            var result = results.First();
            result.ColumnA.Should().Be("a");
            result.ColumnB.Should().Be("b");
            result.NotOnSpreadsheet.Should().BeNull();
        }
        
        [Fact]
        public void AllColumnsMustBeMappedToProperties_True_Works()
        {
            Action action = () =>
            {
                var testSpreadsheetLocation = TestHelper.TestsheetPath("IgnoreColumns.xlsx");
                testSpreadsheetLocation.ExcelToEnumerable<AllColumnsMustBeMappedToPropertiesTestClass>(x => x
                    .AllColumnsMustBeMappedToProperties(true)
                );
            };
            action.Should().ThrowExactly<ExcelToEnumerableInvalidHeaderException>().WithMessage("Missing properties: 'notonclass'.");
        }
        
        [Fact]
        public void AllColumnsMustBeMappedToProperties_Attribute_Works()
        {
            Action action = () =>
            {
                var testSpreadsheetLocation = TestHelper.TestsheetPath("IgnoreColumns.xlsx");
                testSpreadsheetLocation
                    .ExcelToEnumerable<AllColumnsMustBeMappedToPropertiesAttributeTestClass>();
            };
            action.Should().ThrowExactly<ExcelToEnumerableInvalidHeaderException>().WithMessage("Missing properties: 'notonclass'.");
        }
        
        [Fact]
        public void AllColumnsMustBeMappedToPropertiesFalseWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("IgnoreColumns.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<AllColumnsMustBeMappedToPropertiesTestClass>(x => x
                .AllColumnsMustBeMappedToProperties(false)
            ).ToArray();
            results.Length.Should().Be(1);
            var result = results[0];
            result.ColumnA.Should().Be("a");
            result.ColumnB.Should().Be("b");
        }

        [Fact]
        public void NoHeaderWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("NoHeaderTests.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<NoHeaderTestClass>(
                x => x.UsingHeaderNames(false)
            ).ToArray();
            result.First().ColumnA.Should().Be("Value1");
            result.First().ColumnB.Should().Be(1234);
            result.Last().ColumnA.Should().Be("Value2");
            result.Last().ColumnB.Should().Be(3456);
        }
        
        [Fact]
        public void UsingHeaderNames_True_Works()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("NoHeaderTests.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<NoHeaderTestClass>(
                x => x.UsingHeaderNames(true)
            );
            result.Count().Should().Be(0);
        }
        
        [Fact]
        public void UsingHeaderNames_False_Attribute_Works()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("NoHeaderTests.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<NoHeaderAttributeTestClass>().ToArray();
            result.First().ColumnA.Should().Be("Value1");
            result.First().ColumnB.Should().Be(1234);
            result.Last().ColumnA.Should().Be("Value2");
            result.Last().ColumnB.Should().Be(3456);
        }

        [Fact]
        public void NoHeaderWithNumberedColumnsWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("NoHeaderTests.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<NoHeaderTestClass>(
                x => x.UsingHeaderNames(false)
                    .UsingSheet("Numbered Columns")
                    .Property(y => y.ColumnA).MapsToColumnNumber(2)
                    .Property(y => y.ColumnB).MapsToColumnNumber(1)
            ).ToArray();
            result.First().ColumnA.Should().Be("Value1");
            result.First().ColumnB.Should().Be(1234);
            result.Last().ColumnA.Should().Be("Value2");
            result.Last().ColumnB.Should().Be(3456);
        }

        [Fact]
        public void MapsToColumnNumberWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesTest.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesTestClass>(x =>
                x.StartingFromRow(4)
                    .UsingHeaderNames(false)
                    .Property(y => y.ColumnA).MapsToColumnNumber(1)
                    .Property(y => y.ColumnB).MapsToColumnNumber(2)
                    .Property(y => y.ColumnC).MapsToColumnNumber(3)
                    .Property(y => y.ColumnAA).IgnoreColumn()
                    .Property(y => y.IgnoreThisProperty).IgnoreColumn()
            ).ToArray();
            result.Count().Should().Be(2);
            result.First().ColumnA.Should().Be("A");
            result.First().ColumnB.Should().Be("B");
            result.First().ColumnC.Should().Be("C");
            result.Last().ColumnA.Should().Be("Z");
            result.Last().ColumnB.Should().Be("Y");
            result.Last().ColumnC.Should().Be("X");
        }
        
        [Fact]
        public void MapsToColumnNumber_Attribute_Works()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesAttributeTest.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesAttributeTestClass>().ToArray();
            result.Length.Should().Be(2);
            result.First().ColumnA.Should().Be("A");
            result.First().ColumnB.Should().Be("B");
            result.First().ColumnC.Should().Be("C");
            result.Last().ColumnA.Should().Be("Z");
            result.Last().ColumnB.Should().Be("Y");
            result.Last().ColumnC.Should().Be("X");
        }

        [Fact]
        public void MapsToColumnLetterWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesTest.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesTestClass>(x =>
                x.StartingFromRow(4)
                    .UsingHeaderNames(false)
                    .Property(y => y.ColumnA).MapsToColumnLetter("A")
                    .Property(y => y.ColumnB).MapsToColumnLetter("B")
                    .Property(y => y.ColumnC).MapsToColumnLetter("C")
                    .Property(y => y.ColumnAA).MapsToColumnLetter("AA")
                    .Property(y => y.Row).MapsToRowNumber()
                    .Property(y => y.IgnoreThisProperty).IgnoreColumn()
            ).ToArray();
            result.Length.Should().Be(2);
            result.First().ColumnA.Should().Be("A");
            result.First().ColumnB.Should().Be("B");
            result.First().ColumnC.Should().Be("C");
            result.First().Row.Should().Be(4);
            result.First().ColumnAA.Should().Be("XXX");
            result.Last().ColumnA.Should().Be("Z");
            result.Last().ColumnB.Should().Be("Y");
            result.Last().ColumnC.Should().Be("X");
            result.Last().ColumnAA.Should().BeNull();
        }
        
        [Fact]
        public void MapsToColumnLetterAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesAttributeTest.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesColumnLetterAttributeTestClass>().ToArray();
            result.Length.Should().Be(2);
            result.First().ColumnA.Should().Be("A");
            result.First().ColumnB.Should().Be("B");
            result.First().ColumnC.Should().Be("C");
            result.Last().ColumnA.Should().Be("Z");
            result.Last().ColumnB.Should().Be("Y");
            result.Last().ColumnC.Should().Be("X");
        }

        [Fact]
        public void ThrowsConfigExceptionIfNotAllPropertiesMapped()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesTest.xlsx");
            ExcelToEnumerableConfigException exception = null;
            try
            {
                var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesTestClass>(
                    x => x
                        .UsingHeaderNames(false)
                        .StartingFromRow(4)
                        .Property(y => y.ColumnC).MapsToColumnLetter("C")
                );
            }
            catch (ExcelToEnumerableConfigException ex)
            {
                exception = ex;
            }

            exception.Should().NotBeNull();
            exception.Message.Should()
                .Be(
                    "Trying to map property 'ColumnB' to column 'C' but that column is already mapped to property 'ColumnC'.");
        }

        [Fact]
        public void ExceptionValuesDictionaryIsCorrect()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet4Errors.xlsx");
            var exceptionList = new List<Exception>();
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x =>
                    x.OutputExceptionsTo(exceptionList)
                        .UsingSheet("Sheet3ErrorValuesDictionaryTest")
            );
            exceptionList.Count.Should().Be(1);
            var exception = exceptionList.First() as ExcelToEnumerableCellException;
            exception.Should().NotBeNull();
            var rowValues = exception.RowValues;
            rowValues["A"].Should().Be("abc123");
            rowValues["B"].Should().Be("1");
            rowValues["D"].Should().Be("notADecimal");
            rowValues.ContainsKey("C").Should().BeFalse();
            var unmappedObject = exception.UnmappedObject as TestClass;
            unmappedObject.Should().NotBeNull();
            unmappedObject.String.Should().Be("abc123");
            unmappedObject.Int.Should().Be(1);
            unmappedObject.DateTime.Should().BeNull();
            unmappedObject.Decimal.Should().BeNull();
        }

        [Fact]
        public void OptionalColumns1()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var result1 = testSpreadsheetLocation.ExcelToEnumerable<OptionalColumnsTestClass>(x => x
                .UsingSheet("2Columns")
            ).ToArray();
            result1.Length.Should().Be(2);
            var first = result1.First();
            first.Name.Should().Be("Chris");
            first.Fee1.Should().Be((decimal) 1.1);
            first.Fee2.Should().Be(0);
            first.Fee3.Should().Be(0);
            var last = result1.Last();
            last.Name.Should().Be("Adrian");
            last.Fee1.Should().Be((decimal) 2.2);
        }

        [Fact]
        public void OptionalColumns2()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var result1 = testSpreadsheetLocation.ExcelToEnumerable<OptionalColumnsTestClass>(x => x
                .UsingSheet("4Columns")
            ).ToArray();
            result1.Count().Should().Be(2);
            var first = result1.First();
            first.Name.Should().Be("Chris");
            first.Fee1.Should().Be((decimal) 1.1);
            first.Fee2.Should().Be(55);
            first.Fee3.Should().Be(-23);
            var last = result1.Last();
            last.Name.Should().Be("Adrian");
            last.Fee1.Should().Be((decimal) 2.2);
            last.Fee2.Should().Be((decimal) 3.2);
            last.Fee3.Should().Be((decimal) 2.4);
        }

        [Fact]
        public void OptionalColumns3()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var action = new Action(() =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<OptionalColumnsTestClass>(x => x
                        .UsingSheet("2Columns")
                        .AllColumnsMustBeMappedToProperties(true)
                        .Property(y => y.Fee3).OptionalColumn(false) //In this example Fee3 is now a mandatory column
                );
            });
            action.Should().Throw<ExcelToEnumerableInvalidHeaderException>();
        }

        [Fact]
        public void OptionalColumns4()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var result1 = testSpreadsheetLocation.ExcelToEnumerable<OptionalColumnsTestClass>(x => x
                .UsingSheet("4Columns")
                .AllColumnsMustBeMappedToProperties(false)
                .AllPropertiesMustBeMappedToColumns(false)
            ).ToArray();
            result1.Count().Should().Be(2);
            var first = result1.First();
            first.Name.Should().Be("Chris");
            first.Fee1.Should().Be((decimal) 1.1);
            first.Fee2.Should().Be(55);
            first.Fee3.Should().Be(-23);
            var last = result1.Last();
            last.Name.Should().Be("Adrian");
            last.Fee1.Should().Be((decimal) 2.2);
            last.Fee2.Should().Be((decimal) 3.2);
            last.Fee3.Should().Be((decimal) 2.4);
        }

        [Fact]
        public void OptionalColumns5()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var action = new Action(() =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<OptionalColumnsTestClass>(x => x
                    .UsingSheet("2Columns")
                    .AllPropertiesMustBeMappedToColumns(true)
                    .Property(y => y.Fee2).OptionalColumn()
                );
            });
            action.Should().Throw<ExcelToEnumerableInvalidHeaderException>()
                .WithMessage("Missing headers: 'fee3'. ");
        }
        
        [Fact]
        public void OptionalColumnsAttribute()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var action = new Action(() =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<OptionalParametersAttributeTestClass>();
            });
            action.Should().Throw<ExcelToEnumerableInvalidHeaderException>()
                .WithMessage("Missing headers: 'fee3'. ");
        }

        [Fact]
        public void BooleansWork()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("Booleans.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<BooleansTestClass>().ToArray();
            results.Count().Should().Be(2);
            var r1 = results.First();
            r1.BoolAsInt.Should().BeFalse();
            r1.BoolAsText.Should().BeFalse();
            r1.NullableBool.Should().BeTrue();
            r1.MixedFormat.Should().BeTrue();

            var r2 = results.Last();
            r2.BoolAsInt.Should().BeTrue();
            r2.BoolAsText.Should().BeTrue();
            r2.NullableBool.Should().BeNull();
            r2.MixedFormat.Should().BeTrue();
        }

        [Fact]
        public void OnReadHeaderRowWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var headerDictionary = new Dictionary<int, string>();
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>
            (x => x
                .StartingFromRow(4)
                .UsingSheet("HeaderOnRow2")
                .HeaderOnRow(2)
                .OnReadingHeaderRow(y =>
                {
                    foreach (var item in y)
                    {
                        headerDictionary.Add(item.Key, item.Value);
                    }
                }));
            result.Count().Should().Be(3);
            headerDictionary[1].Should().Be("String");
            headerDictionary[2].Should().Be("Int");
            headerDictionary[3].Should().Be("DateTime");
            headerDictionary[4].Should().Be("Decimal");
        }

        [Fact]
        public void CollectionConfigurationWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<CollectionTestClass>(
                x => x.StartingFromRow(1)
                    .UsingSheet("Sheet3CollectionExample")
                    .Property(y => y.Collection).MapFromColumns("CollectionColumn1", "CollectionColumn2")
            ).ToArray();
            var firstResult = result.FirstOrDefault();
            firstResult.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            firstResult.Collection.Should().NotBeNull();
            firstResult.Collection.First().Should().Be("a");
        }
        
        [Fact]
        public void CollectionConfigurationAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<CollectionAttributeTestClass>().ToArray();
            var firstResult = result.FirstOrDefault();
            firstResult.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            firstResult.Collection.Should().NotBeNull();
            firstResult.Collection.First().Should().Be("a");
        }

        [Fact]
        public void HeaderInRow2Works()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x => x
                    .UsingSheet("HeaderOnRow2")
                    .HeaderOnRow(2)
                    .StartingFromRow(4)
            ).ToArray();
            result.Length.Should().Be(3);
            result.First().String.Should().Be("abc123");
            result.Last().String.Should().Be("zxy123");
        }
        
        [Fact]
        public void HeaderInRowAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestHeaderInRowAttributeClass>().ToArray();
            result.Length.Should().Be(3);
            result.First().String.Should().Be("abc123");
            result.Last().String.Should().Be("zxy123");
        }

        [Fact]
        public void CustomColumnNameWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x => x.StartingFromRow(2)
                    .EndingWithRow(2)
                    .UsingSheet("CustomColumnName")
                    .Property(y => y.Int).MapsToColumnNamed("IntCustomName")
            );
            result.First().Int.Should().Be(1);
        }
        
        [Fact]
        public void CustomColumnNameAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<CustomColumnNameAttributeTestClass>();
            result.First().Int.Should().Be(1);
        }

        [Fact]
        public void OptionalColumnWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(2)
                    .Property(y => y.String).IgnoreColumn()
            ).ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be(null);
            row1.Int.Should().Be(1);
            row1.DateTime.Should().Be(new DateTime(2012, 12, 31));
            row1.Decimal.Should().Be(1.234);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be(null);
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
        }
        
        [Fact]
        public void OptionalColumnAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OptionalColumnAttributeTestClass>().ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be(null);
            row1.Int.Should().Be(1);
            row1.DateTime.Should().Be(new DateTime(2012, 12, 31));
            row1.Decimal.Should().Be(1.234);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be(null);
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
        }

        [Fact]
        public void UseCustomMappingWorks()
        {
            var measureReverseLookup = new Dictionary<string, int>
            {
                {"Each", 1},
                {"kg", 2},
                {"g", 3},
                {"lt", 4},
                {"ml", 5}
            };
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MostComplexExample.xlsx");
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleWithCustomMappingTestClass>(
                    x => x.StartingFromRow(2)
                        .EndingWithRow(14)
                        .HeaderOnRow(2)
                        .OutputExceptionsTo(exceptionList)
                        .UsingHeaderNames(true)
                        .UsingSheet("Prices")
                        .Property(y => y.MinimumOrderQuantity).NotNull()
                        .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                        .Property(y => y.Vat).NotNull()
                        .Property(y => y.Vat).ShouldBeOneOf("Standard", "Reduced", "2nd Reduced", "Zero")
                        .Property(y => y.MeasureId).NotNull()
                        .Property(y => y.MeasureId).MapsToColumnNamed("Measure")
                        .Property(y => y.MeasureId).UsesCustomMapping(
                            z =>
                            {
                                if (measureReverseLookup.ContainsKey(z.ToString()))
                                {
                                    return measureReverseLookup[z.ToString()];
                                }

                                throw new KeyNotFoundException();
                            })
                        .Property(y => y.PslCategory).NotNull()
                        .Property(y => y.PslCategory).ShouldBeOneOf("BBB")
                        .Property(y => y.Sku).ShouldBeUnique()
                        .Property(y => y.Sku).NotNull()
                        .Property(y => y.TranslatedSupplierDescriptions).NotNull()
                        .Property(y => y.TranslatedSupplierDescriptions)
                        .MapFromColumns("supplierdescription", "dutchdescription")
                        .Property(y => y.Price).NotNull()
                        .Property(y => y.Price).ShouldBeGreaterThan(0)
                        .Property(y => y.Unit).NotNull()
                        .Property(y => y.Unit).ShouldBeGreaterThan(0)
                        .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403", "depotid-14760")
                ).ToArray();

            var result1 = result.First();
            result1.MeasureId.Should().Be(1);
            var result2 = result.Last();
            result2.MeasureId.Should().Be(1);
        }

        [Fact]
        public void ComplexExampleWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet2.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass2>(
                x => x.StartingFromRow(3)
                    .UsingHeaderNames(false)
            );
            var result1 = result.First();
            result1.Sku.Should().Be("1.222");
            result1.SupplierCategory.Should().Be("Fine foods");
            result1.Store.Should().Be("Fine foods");
            result1.SupplierDescription.Should().Be("XXXX");
            result1.Price.Should().Be(1.1);
            result1.Vat.Should().Be("Zero");
            result1.CaseQty.Should().Be(1);
            result1.PackQty.Should().Be(1);
            result1.Unit.Should().Be(2.6);
            result1.Measure.Should().Be("kg");
            result1.UomDescription.Should().Be("1X2.6KG");
            result1.Origin.Should().BeNull();
            result1.Brand.Should().BeNull();
            result1.LastPeriodVolume.Should().BeNull();
            var result2 = result.Skip(1).First();
            result2.Sku.Should().Be("9.999");
            result.Count().Should().Be(2999);
        }

        [Fact]
        public void ReadsGermanDecimalsCorrectlyWhenCultureSetInThread()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            var testSpreadsheetLocation = TestHelper.TestsheetPath("GermanDecimals.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<GermanDecimals>(x => x
                .Property(y => y.GermanDecimalAsText)
                .UsesCustomMapping(cellValueObject => Convert.ToDouble(cellValueObject)));
            results.Count().Should().Be(2);
            results.First().GermanDecimal.Should().Be(1.234);
            results.First().UkDecimal.Should().Be(1.234);
            results.First().GermanDecimalAsText.Should().Be(1.234);
            results.Last().GermanDecimal.Should().Be(9.876);
            results.Last().UkDecimal.Should().Be(9.876);
            results.Last().GermanDecimalAsText.Should().Be(9.876);
        }

        [Fact]
        public void ReadsGermanDecimalsCorrectlyWhenCultureSetInThreadDifferentSheetFormat()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            var testSpreadsheetLocation = TestHelper.TestsheetPath("doubletest.xlsx");
            var contracts = testSpreadsheetLocation.ExcelToEnumerable<Contract>();
            contracts.Last().Pr01_Jb.Should().Be(289.99);
        }

        [Fact]
        public void ReadsGermanDecimalsCorrectlyWhenCultureSetInCustomMapping()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("GermanDecimals.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<GermanDecimals>(x => x
                .Property(y => y.GermanDecimalAsText).UsesCustomMapping(cellValueObject =>
                    Convert.ToDouble(cellValueObject, new CultureInfo("de-DE")))).ToArray();
            results.Count().Should().Be(2);
            results.First().GermanDecimal.Should().Be(1.234);
            results.First().UkDecimal.Should().Be(1.234);
            results.First().GermanDecimalAsText.Should().Be(1.234);
            results.Last().GermanDecimal.Should().Be(9.876);
            results.Last().UkDecimal.Should().Be(9.876);
            results.Last().GermanDecimalAsText.Should().Be(9.876);
        }

        [Fact]
        public void CreatingBlankEntitiesOptionsWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet2.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass2>(x =>
                x.StartingFromRow(3)
                    .UsingHeaderNames(false)
                    .BlankRowBehaviour(BlankRowBehaviour.CreateEntity)
            );
            result.Count().Should().Be(3000);
        }

        [Fact]
        public void ThrowsCorrectExcelToEnumerableCellExceptionForInvalidDouble()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("Exceptions.xlsx");
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<DoubleTestClass>();
            }
            catch (ExcelToEnumerableCellException e)
            {
                e.Message.Should()
                    .Be("Unable to set value 'Not a double' to property 'double' on row 2 column A. Value is invalid");
                e.InnerException.GetType().Should().Be(typeof(InvalidCastException));
                e.Column.Should().Be("A");
                e.RowNumber.Should().Be(2);
                e.PropertyName.Should().Be("double");
            }
        }

        [Fact]
        public void ReadsExcelOnlineFormatLargeDoubleCorrectly()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("LargeDouble.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<DoubleTestClass>();
            results.First().Double.Should().Be(7060151014090010);
        }

        [Fact]
        public void CustomValidatorWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("CustomValidator.xlsx");
            var exceptionList = new List<Exception>();
            var results = testSpreadsheetLocation.ExcelToEnumerable<CustomValidatorTestClass>(
                x => x.Property(y => y.IsItCheese).UsesCustomValidator(
                        o => o != null && o.ToString().IndexOf("cheese", StringComparison.Ordinal) > -1,
                        "Should contain 'cheese'")
                    .OutputExceptionsTo(exceptionList)
            ).ToArray();
            results.Length.Should().Be(2);
            results.First().IsItCheese.Should().Be("cheese");
            results.Last().IsItCheese.Should().Be("also cheese");
            exceptionList.Count().Should().Be(1);
            exceptionList.First().Message.Should()
                .Be("Unable to set value 'regret' to property 'isitcheese' on row 4 column A. Should contain 'cheese'");
        }

        [Fact(Skip = "Temporarily skipped as slowing tests down. CH 25112020")]
        public void ReadsMillionsOfCells()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OneMillionCells.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<MillionCellTestClass>().ToArray();
            results.Count().Should().Be(99999);
            var i = 1;
            foreach (var result in results)
            {
                result.Col1.Should().Be(7060151014090012 + i);
                result.Col2.Should().Be("X" + (i + 2 + 7060151014090011));
                result.Col3.Should().Be("X" + (i + 3 + 7060151014090011));
                result.Col4.Should().Be("X" + (i + 4 + 7060151014090011));
                result.Col5.Should().Be("X" + (i + 5 + 7060151014090011));
                result.Col6.Should().Be("X" + (i + 6 + 7060151014090011));
                result.Col7.Should().Be("X" + (i + 7 + 7060151014090011));
                result.Col8.Should().Be("X" + (i + 8 + 7060151014090011));
                result.Col9.Should().Be("X" + (i + 9 + 7060151014090011));
                result.Col10.Should().Be("X" + (i + 10 + 7060151014090011));
                i++;
            }
        }

        [Fact]
        public void ThrowsCorrectExcelToEnumerableCellExceptionForCustomMapping()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("Exceptions.xlsx");
            try
            {
                var results = testSpreadsheetLocation.ExcelToEnumerable<DoubleTestClass>(
                    x => x.Property(y => y.Double).UsesCustomMapping(z => { return double.Parse(z.ToString()); }));
            }
            catch (ExcelToEnumerableCellException e)
            {
                e.Message.Should()
                    .Be("Unable to set value 'Not a double' to property 'double' on row 2 column A. Value is invalid");
                e.InnerException.GetType().Should().Be(typeof(FormatException));
                e.Column.Should().Be("A");
                e.RowNumber.Should().Be(2);
                e.PropertyName.Should().Be("double");
            }
        }

        [Fact]
        public void DefaultExceptionHandlingWorks()
        {
            ///CSH 12-11-2019: This should throw an exception on Row 1 as we're expecting the first row 
            ///in the spreadsheet to be data, but in 'TestSpreadsheet4Errors.xlsx' the first row
            ///is a header
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet4Errors.xlsx");
            ExcelToEnumerableCellException excelToEnumerableCellException = null;

            void Action()
            {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.UsingHeaderNames(false));
            }

            try
            {
                Action();
            }
            catch (ExcelToEnumerableCellException e)
            {
                excelToEnumerableCellException = e;
            }

            excelToEnumerableCellException.Should().NotBeNull();
            excelToEnumerableCellException.RowNumber.Should().Be(1);
            excelToEnumerableCellException.Column.Should().Be("B");
            excelToEnumerableCellException.RowValues["A"].Should().Be("String");
            excelToEnumerableCellException.RowValues["B"].Should().Be("Int");
            excelToEnumerableCellException.RowValues["C"].Should().Be("DateTime");
            excelToEnumerableCellException.RowValues["D"].Should().Be("Decimal");
        }

        /// <summary>
        ///     CSH 03122019 - Handles prefixed workbook xml
        /// </summary>
        [Fact]
        public void PrefixedSpreadsheetWorks()
        {
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("PrefixedSpreadsheet.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                x => x.OutputExceptionsTo(exceptionList).UsingHeaderNames(false));
            result.Count().Should().Be(0);
        }

        [Fact]
        public void OfficeGeneratedFileWorks()
        {
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MSOfficeGenerated.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                x => x.OutputExceptionsTo(exceptionList).UsingHeaderNames(false));
            result.Count().Should().Be(0);
        }

        [Fact]
        public void GoogleSheetsGeneratedFileWorks()
        {
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("GoogleSheetsGenerated.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                x => x.OutputExceptionsTo(exceptionList)
                    .UsingHeaderNames(false)
            );
            result.Count().Should().Be(0);
        }

        [Fact]
        public void FromStreamHandlingWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var fileStream = new FileStream(testSpreadsheetLocation, FileMode.Open, FileAccess.Read);
            var result = fileStream.ExcelToEnumerable<TestClass>(x => x.StartingFromRow(2));
            result.Count().Should().Be(3);
        }

        [Fact]
        public void CollectionAndCustomMapperWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var fileStream = new FileStream(testSpreadsheetLocation, FileMode.Open, FileAccess.Read);
            var result = fileStream.ExcelToEnumerable<CollectionTestClass>(x =>
                x
                    .StartingFromRow(1)
                    .UsingSheet("Sheet3CollectionExample")
                    .Property(y => y.Collection).MapFromColumns("CollectionColumn1", "CollectionColumn2")
                    .Property(y => y.Collection).UsesCustomMapping(z => { return z.ToString().ToUpperInvariant(); })
            ).ToArray();
            result.Count().Should().Be(2);
            result.First().Collection.First().Should().Be("A");
        }

        [Fact]
        public void FromStringWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.StartingFromRow(2)).ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123");
            row1.Int.Should().Be(1);
            row1.DateTime.Should().Be(new DateTime(2012, 12, 31));
            row1.Decimal.Should().Be(1.234);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be("zyx987");
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
        }

        [Fact]
        public void InternalSettersWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<InternalSetterTestClass>(x => x.StartingFromRow(2))
                .ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123");
            row1.Int.Should().Be(1);
            row1.DateTime.Should().Be(new DateTime(2012, 12, 31));
            row1.Decimal.Should().Be(1.234);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be("zyx987");
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
        }

        [Fact]
        public void MapRowNumberWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClassWithRowNumber>(
                x => x.StartingFromRow(2)
                    .Property(y => y.Row).MapsToRowNumber()
            ).ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123");
            row1.Int.Should().Be(1);
            row1.DateTime.Should().Be(new DateTime(2012, 12, 31));
            row1.Decimal.Should().Be(1.234);
            row1.Row.Should().Be(2);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be("zyx987");
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            row2.Row.Should().Be(3);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
            row3.Row.Should().Be(4);
        }
        
        [Fact]
        public void MapRowNumberAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<AttributeTestClassWithRowNumber>().ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.Row.Should().Be(2);
            var row2 = result.Skip(1).First();
            row2.Row.Should().Be(3);
            var row3 = result.Skip(2).First();
            row3.Row.Should().Be(4);
        }

        [Fact]
        public void ValidateHeaderWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            ExcelToEnumerableSheetException sheetException = null;
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x
                    .StartingFromRow(1)
                    .AllColumnsMustBeMappedToProperties(true)
                    .UsingSheet("BadHeaderNames")
                    .AggregateExceptions()
                );
            }
            catch (ExcelToEnumerableSheetException e)
            {
                sheetException = e;
            }

            sheetException.Should().NotBeNull();
        }

        [Fact]
        public void EmptyHeaderNames()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            ExcelToEnumerableInvalidHeaderException sheetException = null;
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x
                    .StartingFromRow(1)
                    .AllColumnsMustBeMappedToProperties(true)
                    .UsingSheet("EmptyColumnNames")
                    .AggregateExceptions()
                );
            }
            catch (ExcelToEnumerableInvalidHeaderException e)
            {
                sheetException = e;
            }

            // ReSharper disable once PossibleNullReferenceException
            sheetException.MissingProperties.Count().Should().Be(2);
            // ReSharper disable ReturnValueOfPureMethodIsNotUsed
            sheetException.MissingProperties.Single(x => x == "blankcolumnb");
            sheetException.MissingProperties.Single(x => x == "blankcolumnc");
            // ReSharper restore ReturnValueOfPureMethodIsNotUsed
        }

        [Fact]
        public void MaximumValueWorks()
        {
            var list = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(3)
                    .OutputExceptionsTo(list)
                    .Property(y => y.Int).ShouldBeLessThan(2)
            );
            list.Count.Should().Be(1);
            var firstException = (ExcelToEnumerableCellException) list.First();
            firstException.Column.Should().Be("B");
            firstException.RowNumber.Should().Be(3);
        }
        
        [Fact]
        public void MaximumValueAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            ExcelToEnumerableCellException exception = null;
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<ShouldBeLessThanAttributeTestClass>();
            }
            catch (ExcelToEnumerableCellException e)
            {
                exception = e;
            }

            exception.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            exception.Column.Should().Be("B");
            exception.RowNumber.Should().Be(3);
        }

        [Fact]
        public void PreBuiltOptionsWorks()
        {
            var list = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var optionsBuilder = new ExcelToEnumerableOptionsBuilder<TestClass>();
            optionsBuilder.StartingFromRow(2);
            optionsBuilder.OutputExceptionsTo(list);
            optionsBuilder.Property(x => x.Int).ShouldBeLessThan(2);
            testSpreadsheetLocation.ExcelToEnumerable(optionsBuilder);
            list.Count.Should().Be(1);
            var firstException = (ExcelToEnumerableCellException) list.First();
            firstException.Column.Should().Be("B");
            firstException.RowNumber.Should().Be(3);
        }

        [Fact]
        public void MinimumValueValidationWorks()
        {
            var list = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(2)
                    .OutputExceptionsTo(list)
                    .Property(y => y.Int).ShouldBeGreaterThan(1));
            list.Count.Should().Be(1);
            var firstException = (ExcelToEnumerableCellException) list.First();
            firstException.Column.Should().Be("B");
            firstException.RowNumber.Should().Be(2);
        }
        
        [Fact]
        public void MinimumValueAttributeValidationWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            ExcelToEnumerableCellException exception = null;
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<ShouldBeGreaterThanAttributeTestClass>();
            }
            catch (ExcelToEnumerableCellException e)
            {
                exception = e;
            }

            exception.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            exception.Column.Should().Be("B");
            exception.RowNumber.Should().Be(2);
        }

        [Fact]
        public void OutputExceptionsWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet4Errors.xlsx");
            var list = new List<Exception>();
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x => x.OutputExceptionsTo(list).UsingHeaderNames(false));
            list.Count.Should().Be(3);
        }

        [Fact]
        public void NotNullPropertiesWorks()
        {
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x    
                .StartingFromRow(2)
                .Property(y => y.String).NotNull()
                .Property(y => y.Int).NotNull()
                .Property(y => y.DateTime).NotNull()
                .OutputExceptionsTo(exceptionList));
            result.Count().Should().Be(2);
            exceptionList.Count.Should().Be(3);
            var firstException = (ExcelToEnumerableCellException) exceptionList.First();
            firstException.Column.Should().Be("A");
            firstException.RowNumber.Should().Be(4);
        }
        
        [Fact]
        public void NotNullPropertiesAttributeWorks()
        {
            AggregateException exception = null;
            try
            {
                var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
                testSpreadsheetLocation.ExcelToEnumerable<NotNullAttributeTestClass>();
            }
            catch (AggregateException e)
            {
                exception = e;
            }

            exception.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            var exceptions = exception.InnerExceptions;
            exceptions.Count.Should().Be(3);
            var firstException = (ExcelToEnumerableCellException) exceptions.First();
            firstException.Column.Should().Be("A");
            firstException.RowNumber.Should().Be(4);
        }

        [Fact]
        public void NotNullColumnAndAggregateExceptionWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            AggregateException testException = null;
            try
            {
                var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                    x.StartingFromRow(2)
                        .AggregateExceptions()
                        .Property(y => y.String).NotNull()
                        .Property(y => y.Int).NotNull()
                        .Property(y => y.DateTime).NotNull()
                );
            }
            catch (AggregateException e)
            {
                testException = e;
            }

            testException.InnerExceptions.Count().Should().Be(3);
            var firstException = (ExcelToEnumerableCellException) testException.InnerExceptions.First();
            firstException.Column.Should().Be("A");
            firstException.RowNumber.Should().Be(4);
        }

        [Fact]
        public void SelectWorksheetByNameWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x => x.StartingFromRow(2)
                    .UsingSheet("Sheet2NonDefaultName")
            );
            result.Count().Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123sheet2");
        }

        [Fact]
        public void SelectWorksheetByNumberWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x => x.StartingFromRow(2)
                    .UsingSheet(1)
            );
            result.Count().Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123sheet2");
        }
        
        [Fact]
        public void SelectWorksheetByNumberAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<WorksheetByNumberAttributeTestClass>().ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123sheet2");
        }

        [Fact]
        public void ShouldBeOneOfValidationWorks()
        {
            var list = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(2)
                    .OutputExceptionsTo(list)
                    .Property(y => y.String)
                    .ShouldBeOneOf("abc123", "Steve", "Steven"));
            list.Count.Should().Be(1);
            var firstException = (ExcelToEnumerableCellException) list.First();
            firstException.Column.Should().Be("A");
            firstException.RowNumber.Should().Be(3);
        }
        
        [Fact]
        public void ShouldBeOneOfAttributeValidationWorks()
        {
            ExcelToEnumerableCellException exception = null;
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<ShouldBeOneOfAttributeTestClass>();
            }
            catch (ExcelToEnumerableCellException e)
            {
                exception = e;
            }

            exception.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            exception.Column.Should().Be("A");
            exception.RowNumber.Should().Be(3);
        }

        [Fact]
        public void ShouldBeOneOfValidationWorksWithCorrectRowNumbers()
        {
            var list = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(2)
                    .OutputExceptionsTo(list)
                    .Property(y => y.String)
                    .ShouldBeOneOf("abc123", "Steve", "Steven"));
            list.Count.Should().Be(1);
            var firstException = (ExcelToEnumerableCellException) list.First();
            firstException.Column.Should().Be("A");
            firstException.RowNumber.Should().Be(3);
        }

        [Fact]
        public void StopReadingOnBlankRowWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet2.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass2>(x =>
                x.StartingFromRow(3)
                    .BlankRowBehaviour(BlankRowBehaviour.StopReading).UsingHeaderNames(false)
            );
            result.Count().Should().Be(1);
        }
        
        [Fact]
        public void BlankRowBehaviourAttributeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet2.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<BlankRowBehaviourAttributeTest>();
            result.Count().Should().Be(1);
        }

        [Fact]
        public void ThrowExceptionOnBlankRowOptionsWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet2.xlsx");
            Action action = () =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass2>(x =>
                    x.StartingFromRow(3).UsingHeaderNames(false)
                        .BlankRowBehaviour(BlankRowBehaviour.ThrowException)
                );
            };
            action.Should().Throw<ExcelToEnumerableRowException>();
        }

        [Fact]
        public void UsingHeaderWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet3DifferentColumnOrder.xlsx");
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.StartingFromRow(1).UsingHeaderNames(true))
                    .ToArray();
            result.Length.Should().Be(3);
            var row1 = result.First();
            row1.String.Should().Be("abc123");
            row1.Int.Should().Be(1);
            row1.DateTime.Should().Be(new DateTime(2012, 12, 31));
            row1.Decimal.Should().Be(1.234);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be("zyx987");
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
        }

        [Fact]
        public void ShouldBeUniqueWorks()
        {
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(2)
                    .UsingSheet("Sheet4DuplicateStringValue")
                    .Property(y => y.String).ShouldBeUnique()
                    .Property(y => y.Int).ShouldBeUnique()
                    .OutputExceptionsTo(exceptionList)
            );
            exceptionList.Count.Should().Be(2);
            result.Count().Should().Be(1);
        }
        
        [Fact]
        public void ShouldBeUniqueAttributeWorks()
        {
            AggregateException exception = null;
            try
            {
                var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
                testSpreadsheetLocation.ExcelToEnumerable<UniqueAttributeTestClass>();
            }
            catch (AggregateException e)
            {
                exception = e;
            }

            exception.Should().NotBeNull();
            // ReSharper disable once PossibleNullReferenceException
            var exceptionList = exception.InnerExceptions;
            exceptionList.Count.Should().Be(2);
        }

        [Fact]
        public void InvalidCastReturnsCorrectColumnNumber()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MostComplexExample.xlsx");
            var exceptionList = new List<Exception>();
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                    x => x.StartingFromRow(8)
                        .EndingWithRow(8)
                        .HeaderOnRow(2)
                        .UsingSheet("Prices")
                        .OutputExceptionsTo(exceptionList)
                        .Property(y => y.MinimumOrderQuantity).NotNull()
                        .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                        .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("dutchdescription")
                        .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403", "depotid-14760")
                );
            result.Count().Should().Be(0);
            var exceptions = exceptionList.Cast<ExcelToEnumerableCellException>().OrderBy(x => x.RowNumber);
            exceptions.Count().Should().Be(1);

            var firstException = exceptions.First();
            firstException.Column.Should().Be("H");
            firstException.RowNumber.Should().Be(8);
        }

        [Fact]
        public void MinimumAndRequiredWorkOkTogether()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MostComplexExample.xlsx");
            var exceptionList = new List<Exception>();
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                    x => x.StartingFromRow(7)
                        .EndingWithRow(8)
                        .HeaderOnRow(2)
                        .UsingSheet("Prices")
                        .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("dutchdescription")
                        .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403", "depotid-14760")
                        .OutputExceptionsTo(exceptionList)
                        .Property(y => y.MinimumOrderQuantity).NotNull()
                        .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                );
            result.Count().Should().Be(0);
            var exceptions = exceptionList.Cast<ExcelToEnumerableCellException>().OrderBy(x => x.RowNumber);
            exceptions.Count().Should().Be(2);

            var firstException = exceptions.First();
            firstException.Column.Should().Be("H");
            firstException.RowNumber.Should().Be(7);
            var secondException = exceptions.Last();
            secondException.Column.Should().Be("H");
            secondException.RowNumber.Should().Be(8);
        }

        [Fact]
        public void DecimalToIntThrowsException()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MostComplexExample.xlsx");
            var exceptionList = new List<Exception>();
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                    x => x
                        .UsingHeaderNames(true)
                        .HeaderOnRow(2)
                        .UsingSheet("Prices")
                        .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("dutchdescription")
                        .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403", "depotid-14760")
                        .StartingFromRow(21)
                        .EndingWithRow(21)
                        .OutputExceptionsTo(exceptionList));
            result.Count().Should().Be(0);
            exceptionList.Count.Should().Be(1);
            var exception = (ExcelToEnumerableCellException) exceptionList.First();
            exception.RowNumber.Should().Be(21);
            exception.Column.Should().Be("J");
        }

        [Fact]
        public void CastToNullableBoolWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<BoolTestClass>(x => x
                .UsingSheet("Sheet5Booleans")
                .UsingHeaderNames(true)
            );
            var row1 = result.First();
            var row2 = result.Skip(1).First();
            var row3 = result.Skip(2).First();
            row1.Bool.Should().BeFalse();
            row1.NullableBool.Should().BeNull();
            row2.Bool.Should().BeTrue();
            row2.NullableBool.Should().BeTrue();
            row3.Bool.Should().BeFalse();
            row3.NullableBool.Should().BeFalse();
        }

        [Fact]
        public void CastToNullableBoolCollectionWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<BoolCollectionTestClass>(x => x
                .UsingSheet("Sheet5Booleans")
                .UsingHeaderNames(true)
                .Property(y => y.BoolCollection).MapFromColumns("Bool", "NullableBool")
            ).ToArray();
            var row1 = result.First();
            var row2 = result.Skip(1).First();
            var row3 = result.Skip(2).First();
            row1.BoolCollection.Should().BeNull();
            row2.BoolCollection.First().Should().BeTrue();
            row2.BoolCollection.First().Should().BeTrue();
            row3.BoolCollection.First().Should().BeFalse();
            row3.BoolCollection.First().Should().BeFalse();
        }


        [Fact]
        public void MappingToDictionaryWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<DictionaryCollectionTestClass>(
                x => x.StartingFromRow(1)
                    .UsingSheet("Sheet3CollectionExample")
                    .UsingHeaderNames(true)
                    .Property(y => y.Collection)
                    .MapFromColumns("CollectionColumn1", "CollectionColumn2")
            );
            var firstResult = result.FirstOrDefault();
            firstResult.Should().NotBeNull();
            firstResult.Collection.Should().NotBeNull();
            firstResult.Collection.First().Key.Should().Be("CollectionColumn1");
            firstResult.Collection.First().Value.Should().Be("a");
        }

        [Fact]
        public void AutoLocateHeaderWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("AutoLocateHeaderTest.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<AutoLocateHeaderTestClass>().ToArray();
            results.Length.Should().Be(1);
            var result = results.First();
            result.Header.Should().Be("A");
            result.Starts.Should().Be(1);
            result.Here.Should().Be(new DateTime(2020, 11, 25));
        }

        [Fact]
        public void FuzzyMatchingOfHeaderNamesWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("FuzzyMatchHeaderNames.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<FuzzyMatchHeaderNamesTestClass>(
                x => x
                    .AllColumnsMustBeMappedToProperties(false)
                    .AllPropertiesMustBeMappedToColumns(false)).ToArray();
            results.Length.Should().Be(1);
            var result = results.First();
            result.FlatCase.Should().Be(1);
            result.UpperFlatCase.Should().Be(1);
            result.CamelCase.Should().Be(1);
            result.PascalCase.Should().Be(1);
            result.SnakeCase.Should().Be(1);
            result.ScreamingSnakeCase.Should().Be(1);
            result.CamelSnakeCase.Should().Be(1);
            result.KebabCase.Should().Be(1);
            result.CobolCase.Should().Be(1);
            result.TrainCase.Should().Be(1);
            result.WhiteSpace.Should().Be(1);
        }

        [Fact]
        public void ColumnsMappedByNumberButAlsoWithUnmappedColumnsWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("ColumnsMappedByNumberButWithUnmappedColumns.xlsx");
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ColumnsMappedByNumberButWithUnmappedColumnsTestClass>(x => x
                    .StartingFromRow(2)
                    .UsingHeaderNames(false)
                    .Property(y => y.PropertyA).MapsToColumnNumber(1)
                    .Property(y => y.PropertyB).MapsToColumnNumber(2)
                ).ToArray();
            result.Length.Should().Be(1);
            result[0].PropertyA.Should().Be("A");
            result[0].PropertyB.Should().Be("B");
        }

        [Fact]
        public void AllNumericFormatsTest()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("AllNumericFormats.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<AllNumericFormatsTestClass>().ToArray();
            result.Length.Should().Be(2);
            result[0].Int.Should().Be(1);
            result[0].Long.Should().Be(4294967296);
            result[0].Double.Should().Be(4294967296);
            result[0].Decimal.Should().Be(23.0123456789012M);
            result[1].Int.Should().Be(-1);
            result[1].Long.Should().Be(-4294967296);
            result[1].Double.Should().Be(-4294967296);
            result[1].Decimal.Should().Be(-23.0123456789012M);
        }

        [Fact]
        public void EndingWithRowNegativeWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("EndingWithRowNegative.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<EndingWithRowNegativeTestClass>(x => x
                .EndingWithRow(-2)).ToArray();
            result.Length.Should().Be(5);
            result[0].Column.Should().Be(1);
            result[1].Column.Should().Be(2);
            result[2].Column.Should().Be(3);
            result[3].Column.Should().Be(4);
            result[4].Column.Should().Be(5);
        }
        
        [Fact]
        public void EndingWithRowNegativeAndNoDimensionWorksheetThrowsNotImplementedException()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("EndingWithRowNegativeNoDimension.xlsx");
            Action action = () =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<EndingWithRowNegativeTestClass>(x => x
                    .EndingWithRow(-2));
            };
            action.Should().ThrowExactly<NotImplementedException>().WithMessage("ExcelToEnumerable is currently unable to handle a negative EndingWithRow value for a worksheet that does not have a WorksheetDimension xml element.");
        }

        [Fact]
        public void RelaxedNumberMatchingWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("FuzzyNumericMatch.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<AllNumericFormatsTestClass>(x => x.RelaxedNumberMatching(true))
                .ToArray();
            result.Length.Should().Be(2);
            
            result[0].Int.Should().Be(1);
            result[0].Long.Should().Be(-123456789);
            result[0].Double.Should().Be(-123456789.222);
            result[0].Decimal.Should().Be(-123456789.222M);
            result[0].NullableInt.Should().Be(1);
            result[0].NullableLong.Should().Be(-123456789);
            result[0].NullableDouble.Should().Be(-123456789.222);
            result[0].NullableDecimal.Should().Be(-123456789.222M);
            
            result[1].Int.Should().Be(1);
            result[1].Long.Should().Be(-123456789);
            result[1].Double.Should().Be(-123456789.222);
            result[1].Decimal.Should().Be(-123456789.222M);
            result[1].NullableInt.Should().Be(null);
            result[1].NullableLong.Should().Be(null);
            result[1].NullableDouble.Should().Be(null);
            result[1].NullableDecimal.Should().Be(null);
        }

        [Fact]
        public void RequiredColumnWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("RequiredColumns.xlsx");
            Action action = () =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<RequiredColumnsTestClass>(x => x
                    .Property(y => y.RequiredProperty).RequiredColumn());
            };
            // ReSharper disable once StringLiteralTypo
            action.Should().ThrowExactly<ExcelToEnumerableInvalidHeaderException>().WithMessage("Missing headers: 'requiredproperty'. ");
        }

        [Fact]
        public void RequiredColumnWorksWithExplicitlyNameColumn()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("RequiredColumnsExplicitlyNamed.xlsx");
            Action action = () =>
            {
                testSpreadsheetLocation.ExcelToEnumerable<RequiredColumnsTestClass>(x => x
                    .Property(y => y.RequiredProperty).RequiredColumn()
                    .Property(y => y.RequiredProperty).MapsToColumnNamed("This column is not on the spreadsheet"));
            };
            // ReSharper disable once StringLiteralTypo
            action.Should().ThrowExactly<ExcelToEnumerableInvalidHeaderException>().WithMessage("Missing headers: 'thiscolumnisnotonthespreadsheet'. ");
        }

        [Fact]
        public void MapNullToNonNullableMapsToDefaultValue()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MapNullToNonNullableProperty.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<MapNullToNonNullablePropertyThrowsExceptionTestClass>()
                .ToArray();
            results.Length.Should().Be(2);
            results[0].Nullable.Should().BeNull();
            results[0].NotNullable.Should().Be(0);
            results[0].String.Should().Be("a");
            results[1].Nullable.Should().Be(2);
            results[1].NotNullable.Should().Be(1);
            results[1].String.Should().Be("b");
        }
        
        [Fact]
        public void ParsesEnumsCorrectly()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("EnumsTestSmall.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<EnumsTestClass>().ToArray();
            results.Length.Should().Be(1);
        }
        
        [Fact]
        public void ParsesEnumsFromJustStringsCorrectly()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("EnumsTestFromStrings.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<StringsEnumsTestClass>().ToArray();
            results.Length.Should().Be(4);
            results[0].ParseFromStrings.Should().Be(StringsEnumsTestClass.ParseFromStringsEnum.Value1);
            results[1].ParseFromStrings.Should().Be(StringsEnumsTestClass.ParseFromStringsEnum.Value1);
            results[2].ParseFromStrings.Should().Be(StringsEnumsTestClass.ParseFromStringsEnum.Value2);
            results[3].ParseFromStrings.Should().Be(StringsEnumsTestClass.ParseFromStringsEnum.Value2);
        }

        [Fact]
        public void ParsesEnumsFromStringsCorrectly()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("EnumsTest.xlsx");
            var exceptionList = new List<Exception>();
            var results = testSpreadsheetLocation.ExcelToEnumerable<EnumsTestClass>(x => x
                .OutputExceptionsTo(exceptionList)
            ).ToArray();
            
            results.Length.Should().Be(7);
            exceptionList.Count.Should().Be(2);

            results[0].ParseFromStrings.Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            results[0].ParseFromInts.Should().Be(EnumsTestClass.ParseFromIntsEnum.Ten);
            
            results[1].ParseFromStrings.Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            results[1].ParseFromInts.Should().Be(EnumsTestClass.ParseFromIntsEnum.Ten);
            
            results[2].ParseFromStrings.Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            results[2].ParseFromInts.Should().Be(EnumsTestClass.ParseFromIntsEnum.Ten);
            
            results[3].ParseFromStrings.Should().Be(EnumsTestClass.ParseFromStringsEnum.Value2);
            results[3].ParseFromInts.Should().Be(EnumsTestClass.ParseFromIntsEnum.Ten);
            
            results[4].ParseFromStrings.Should().Be(0);
            results[4].ParseFromInts.Should().Be(EnumsTestClass.ParseFromIntsEnum.Ten);
            
            results[5].ParseFromStrings.Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            results[5].ParseFromInts.Should().Be(EnumsTestClass.ParseFromIntsEnum.Twenty);
            
            results[6].ParseFromStrings.Should().Be(EnumsTestClass.ParseFromStringsEnum.Value1);
            results[6].ParseFromInts.Should().Be(null);
        }
        
        [Fact]
        public void RowWithInvalidCellStillCreatesEntity()
        {
            var exceptionList = new List<Exception>();
            var results = TestHelper.TestsheetPath("InvalidRow.xlsx").ExcelToEnumerable<InvalidRowTestClass>(o => o
                .OutputExceptionsTo(exceptionList)
                .BlankRowBehaviour(BlankRowBehaviour.CreateEntity)
                );
            results.Count().Should().Be(2);
        }
    }
}