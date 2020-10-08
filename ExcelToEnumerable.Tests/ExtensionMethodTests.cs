using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToEnumerable.Exceptions;
using FluentAssertions;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class ExtensionMethodTests
    {
        private ComplexExampleWithCustomMappingTestClass result2;

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
            action.Should().Throw<AggregateException>().And.InnerExceptions.Count().Should().Be(3);
        }

        [Fact]
        public void IgnorePropertyNamesWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("PropertyNamesIgnoreAndWhitespaceTests.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<IgnorePropertyNamesTestClass>(x => x
                .IgnoreColumsWithoutMatchingProperties()
                .Property(y => y.NotOnSpreadsheet).Ignore()
                .Property(y => y.ColumnA).MapFromColumns("A", "B")
                .Property(y => y.ColumnA).UsesColumnNamed("Column A")
            );
            result.First().ColumnA.Should().Be("a");
            result.First().ColumnB.Should().Be("b");
        }

        [Fact]
        public void NoHeaderWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("NoHeaderTests.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<NoHeaderTestClass>(
                x => x.UsingHeaderNames(false)
            );
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
                    .Property(y => y.ColumnA).UsesColumnNumber(2)
                    .Property(y => y.ColumnB).UsesColumnNumber(1)
            );
            result.First().ColumnA.Should().Be("Value1");
            result.First().ColumnB.Should().Be(1234);
            result.Last().ColumnA.Should().Be("Value2");
            result.Last().ColumnB.Should().Be(3456);
        }

        [Fact]
        public void UsesColumnNumberWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesTest.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesTestClass>(x =>
                x.StartingFromRow(4)
                    .UsingHeaderNames(false)
                    .Property(y => y.ColumnA).UsesColumnNumber(1)
                    .Property(y => y.ColumnB).UsesColumnNumber(2)
                    .Property(y => y.ColumnC).UsesColumnNumber(3)
                    .Property(y => y.ColumnAA).Ignore()
                    .Property(y => y.IgnoreThisProperty).Ignore()
            );
            result.Count().Should().Be(2);
            result.First().ColumnA.Should().Be("A");
            result.First().ColumnB.Should().Be("B");
            result.First().ColumnC.Should().Be("C");
            result.Last().ColumnA.Should().Be("Z");
            result.Last().ColumnB.Should().Be("Y");
            result.Last().ColumnC.Should().Be("X");
        }
        
        [Fact]
        public void UsesColumnLetterWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OrdinalPropertiesTest.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<OrdinalPropertiesTestClass>(x =>
                x.StartingFromRow(4)
                    .UsingHeaderNames(false)
                    .Property(y => y.ColumnA).UsesColumnLetter("A")
                    .Property(y => y.ColumnB).UsesColumnLetter("B")
                    .Property(y => y.ColumnC).UsesColumnLetter("C")
                    .Property(y => y.ColumnAA).UsesColumnLetter("AA")
                    .Property(y => y.Row).MapsToRowNumber()
                    .Property(y => y.IgnoreThisProperty).Ignore()
            );
            result.Count().Should().Be(2);
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
                        .Property(y => y.ColumnC).UsesColumnLetter("C")
                );
            }
            catch (ExcelToEnumerableConfigException ex)
            {
                exception = ex;
            }

            exception.Should().NotBeNull();
            exception.Message.Should().Be("Trying to map property 'ColumnB' to column 'C' but that column is already mapped to property 'ColumnC'. If you're not using header names then all properties need to be mapped to a column or explicitly ignored.");
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
            var result1 = testSpreadsheetLocation.ExcelToEnumerable<OptionalParametersTestClass>(x => x
                .UsingSheet("2Columns")
                .IgnoreColumsWithoutMatchingProperties()
                .Property(y => y.Fee1).Optional()
                .Property(y => y.Fee2).Optional()
                .Property(y => y.Fee3).Optional()
            );
            result1.Count().Should().Be(2);
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
            var result1 = testSpreadsheetLocation.ExcelToEnumerable<OptionalParametersTestClass>(x => x
                .IgnoreColumsWithoutMatchingProperties()
                .UsingSheet("4Columns")
                .Property(y => y.Fee1).Optional()
                .Property(y => y.Fee2).Optional()
                .Property(y => y.Fee3).Optional()
            );
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
                testSpreadsheetLocation.ExcelToEnumerable<OptionalParametersTestClass>(x => x
                        .UsingSheet("2Columns")
                        .IgnoreColumsWithoutMatchingProperties()
                        .Property(y => y.Fee1).Optional()
                        .Property(y => y.Fee2).Optional()
                        .Property(y => y.Fee3) //In this example Fee3 is now a mandatory column
                );
            });
            action.Should().Throw<ExcelToEnumerableInvalidHeaderException>();
        }
        
        [Fact]
        public void OptionalColumns4()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("OptionalColumns.xlsx");
            var result1 = testSpreadsheetLocation.ExcelToEnumerable<OptionalParametersTestClass>(x => x
                .UsingSheet("4Columns")
                .IgnoreColumsWithoutMatchingProperties()
                .AllPropertiesOptionalByDefault()
            );
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
                testSpreadsheetLocation.ExcelToEnumerable<OptionalParametersTestClass>(x => x
                        .UsingSheet("2Columns")
                        .IgnoreColumsWithoutMatchingProperties()
                        .AllPropertiesOptionalByDefault()
                        .Property(y => y.Fee3).Optional()
                        .Property(y => y.Fee3).Optional(false)
                    );
            });
            action.Should().Throw<ExcelToEnumerableInvalidHeaderException>(
                because: "Property Fee3 is not optional and the spreadsheet 2Columns does not contain it");
        }

        [Fact]
        public void BooleansWork()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("Booleans.xlsx");
            var results = testSpreadsheetLocation.ExcelToEnumerable<BooleansTestClass>();
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
                    .UsingHeaderNames(true)
                    .Property(y => y.Collection)
                    .MapFromColumns("CollectionColumn1", "CollectionColumn2")
            );
            var firstResult = result.FirstOrDefault();
            firstResult.Should().NotBeNull();
            firstResult.Collection.Should().NotBeNull();
            firstResult.Collection.First().Should().Be("a");
        }

        [Fact]
        public void HeaderInRow2Works()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(
                x => x.UsingSheet("HeaderOnRow2")
                    .UsingHeaderNames(true)
                    .HeaderOnRow(2)
                    .StartingFromRow(4)
            );
            result.Count().Should().Be(3);
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
                    .UsingHeaderNames(true)
                    .Property(y => y.Int).IsRequired()
                    .Property(y => y.Int).UsesColumnNamed("IntCustomName")
            );
            result.First().Int.Should().Be(1);
        }

        [Fact]
        public void OptionalColumnWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => 
                x.StartingFromRow(2)
                    .Property(y => y.String).Ignore()
                );
            result.Count().Should().Be(3);
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
            var measureReverseLookup = new Dictionary<string, int> {
                { "Each", 1},
                { "kg", 2},
                { "g", 3},
                { "lt", 4},
                { "ml", 5}
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
                        .Property(y => y.MinimumOrderQuantity).IsRequired()
                        .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                        .Property(y => y.Vat).IsRequired()
                        .Property(y => y.Vat).ShouldBeOneOf("Standard", "Reduced", "2nd Reduced", "Zero")
                        .Property(y => y.MeasureId).IsRequired()
                        .Property(y => y.MeasureId).UsesColumnNamed("Measure")
                        .Property(y => y.MeasureId).UsesCustomMapping(
                            (z) => { if (measureReverseLookup.ContainsKey(z.ToString())){
                                return measureReverseLookup[z.ToString()];
                            } throw new KeyNotFoundException(); })
                        .Property(y => y.PslCategory).IsRequired()
                        .Property(y => y.PslCategory).ShouldBeOneOf("BBB")
                        .Property(y => y.Sku).ShouldBeUnique()
                        .Property(y => y.Sku).IsRequired()
                        .Property(y => y.TranslatedSupplierDescriptions).IsRequired()
                        .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("supplierdescription", "dutchdescription")
                        .Property(y => y.Price).IsRequired()
                        .Property(y => y.Price).ShouldBeGreaterThan(0)
                        .Property(y => y.Unit).IsRequired()
                        .Property(y => y.Unit).ShouldBeGreaterThan(0)
                        .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403", "depotid-14760")
                ).ToArray();

            var result1 = result.First();
            result1.MeasureId.Should().Be(1);
            result2 = result.Last();
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
        public void DefaultExceptionHandlingWorks()
        {
            ///CSH 12-11-2019: This should throw an exception on Row 1 as we're expecting the first row 
            ///in the spreadsheet to be data, and in 'TestSpreadsheet4Errors.xlsx' the first row
            ///is a header
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet4Errors.xlsx");
            ExcelToEnumerableCellException excelToEnumerableCellException = null;
            Action action = () => {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.UsingHeaderNames(false));
            };
            try
            {
                action();
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
        /// CSH 03122019 - Handles prefixed workbook xml
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
            .UsingHeaderNames(true)
            .UsingSheet("Sheet3CollectionExample")
            .Property(y => y.Collection).MapFromColumns("CollectionColumn1","CollectionColumn2")
            .Property(y => y.Collection).UsesCustomMapping((z) => { return z.ToString().ToUpperInvariant(); })
            );
            result.Count().Should().Be(2);
            result.First().Collection.First().Should().Be("A");
        }

        [Fact]
        public void FromStringWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.StartingFromRow(2));
            result.Count().Should().Be(3);
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
            var result = testSpreadsheetLocation.ExcelToEnumerable<InternalSetterTestClass>(x => x.StartingFromRow(2));
            result.Count().Should().Be(3);
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
                );
            result.Count().Should().Be(3);
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
        public void ValidateHeaderWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            ExcelToEnumerableSheetException sheetException = null;
            try
            {
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(1)
                .UsingSheet("BadHeaderNames")
                .UsingHeaderNames(true)
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
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(1)
                .UsingSheet("EmptyColumnNames")
                .UsingHeaderNames(true)
                .AggregateExceptions()
                );
            }
            catch (ExcelToEnumerableInvalidHeaderException e)
            {
                sheetException = e;
            }
            sheetException.MissingProperties.Count().Should().Be(2);
            sheetException.MissingProperties.Single(x => x == "**blank column** (b)");
            sheetException.MissingProperties.Single(x => x == "**blank column** (c)");
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
            var firstException = (ExcelToEnumerableCellException)list.First();
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
        public void OutputExceptionsWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet4Errors.xlsx");
            var list = new List<Exception>();
            testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.OutputExceptionsTo(list).UsingHeaderNames(false));
            list.Count().Should().Be(3);
        }

        [Fact]
        public void RequiredColumnWorks()
        {
            var exceptionList = new List<Exception>();
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                x.StartingFromRow(2)
                .UsingHeaderNames(true)
                .Property(y => y.String).IsRequired()
                .Property(y => y.Int).IsRequired()
                .Property(y => y.DateTime).IsRequired()
                .OutputExceptionsTo(exceptionList));
            result.Count().Should().Be(2);
            exceptionList.Count.Should().Be(3);
            var firstException = (ExcelToEnumerableCellException) exceptionList.First();
            firstException.Column.Should().Be("A");
            firstException.RowNumber.Should().Be(4);
        }

        [Fact]
        public void RequiredColumnAndAggregateExceptionWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            AggregateException testException = null;
            try
            {
                var result = testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x =>
                  x.StartingFromRow(2)
                  .AggregateExceptions()
                  .Property(y => y.String).IsRequired()
                  .Property(y => y.Int).IsRequired()
                  .Property(y => y.DateTime).IsRequired()
                  );
            }
            catch (AggregateException e)
            {
                testException = e;
            }
            testException.InnerExceptions.Count().Should().Be(3);
            var firstException = (ExcelToEnumerableCellException)testException.InnerExceptions.First();
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
            var firstException = (ExcelToEnumerableCellException)list.First();
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
                testSpreadsheetLocation.ExcelToEnumerable<TestClass>(x => x.StartingFromRow(1).UsingHeaderNames(true));
            result.Count().Should().Be(3);
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
            exceptionList.Count().Should().Be(2);
            result.Count().Should().Be(1);
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
                    .Property(y => y.MinimumOrderQuantity).IsRequired()
                    .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                    .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("dutchdescription")
                    .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403","depotid-14760")
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
                    .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403","depotid-14760")
                    .OutputExceptionsTo(exceptionList)
                    .Property(y => y.MinimumOrderQuantity).IsRequired()
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
                        .Property(y => y.DepotExclusions).MapFromColumns("depotid-14403","depotid-14760")
                        .StartingFromRow(21)
                        .EndingWithRow(21)
                        .OutputExceptionsTo(exceptionList));
            result.Count().Should().Be(0);
            exceptionList.Count.Should().Be(1);
            var exception = (ExcelToEnumerableCellException)exceptionList.First();
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
            );
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

        /// <summary>
        /// CSH 11012019 These are not real tests, it's just somewhere to check that the code in the documentation example actually compiles
        /// </summary>
        public void Test()
        {
            var exceptionList = new List<Exception>();
            var result = "filePath".ExcelToEnumerable<Product>(
                //Specify the spreadsheet name:
                x => x.UsingSheet("Prices") 
                    
                    //Header on different row:
                    .HeaderOnRow(2) 
                    
                    //Data on different row:
                    .StartingFromRow(4) 
                    
                    //Stop reading data on row 10:
                    .EndingWithRow(10)
                    
                    //Add validation exceptions to a list instead of throwing them:
                    .OutputExceptionsTo(exceptionList)
                    
                    //Create a Product class even if the spreadsheet row is blank:
                    .BlankRowBehaviour(BlankRowBehaviour.CreateEntity)
                    
                    //Execute arbitrary code after the header row has been read:
                    .OnReadingHeaderRow(headerRowValues => {
                        foreach (var item in headerRowValues)
                        {
                            Console.WriteLine($"Found header label {item.Value} on column {item.Key}");
                        }
                    })
                    
                    //Lots of property validation options:
                    .Property(y => y.ShippingLabel).Ignore()
                    .Property(y => y.MinimumOrderQuantity).IsRequired()
                    .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                    .Property(y => y.Vat).ShouldBeOneOf("Standard", "Reduced", "2nd Reduced", "Zero")
                    .Property(y => y.Measure).IsRequired()
                    .Property(y => y.Id).ShouldBeUnique()
                    .Property(y => y.SupplierDescription).UsesColumnNamed("Supplier Description")
                    
                    //Map a collection to a set of columns:
                    .Property(y => y.SalesVolumes).MapFromColumns("Jan-Mar","Apr-Jun","Jul-Sep","Oct-Dec")
                    
                    //Use a custom mapper to map to a Boolean property, for example:
                    .Property(y => y.IsOnPromotion).UsesCustomMapping(cellValueObject => cellValueObject.ToString() == "Yes!")
                    
                    //Map the sheet row number to a property on a class:
                    .Property(y => y.SpreadsheetRowNumber).MapsToRowNumber()
            );
        }
        
        public class SpreadsheetRow
        {
            public string Store {get;set;}
            public string Area {get;set;}
            public string MonFriTimes {get;set;}
            public string SatTimes {get;set;}
            public string SunTimes {get;set;}
            
            public int RowNumber {get;set;}

        }
        
        public void NoValidationConfigExample()
        {
            var spreadsheetStream = new MemoryStream();
            var spreadsheetData = spreadsheetStream.ExcelToEnumerable<SpreadsheetRow>(o => o
                .UsingHeaderNames(false)
                .StartingFromRow(3) // - The data in this example starts on the 3rd row
                .Property(c => c.Store).UsesColumnLetter("A")
                .Property(c => c.Area).UsesColumnLetter("B")
                .Property(c => c.MonFriTimes).UsesColumnLetter("C")
                .Property(c => c.SatTimes).UsesColumnLetter("D")
                .Property(c => c.SunTimes).UsesColumnLetter("E")
                .Property(c => c.RowNumber).MapsToRowNumber());
        }
    }

}