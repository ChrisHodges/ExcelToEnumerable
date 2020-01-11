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
                        .Property(y => y.PslCategory).ShouldBeOneOf("Frozen Foods")
                        .Property(y => y.Sku).ShouldBeUnique()
                        .Property(y => y.Sku).IsRequired()
                        .Property(y => y.TranslatedSupplierDescriptions).IsRequired()
                        .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("Supplier Description", "Dutch Description")
                        .Property(y => y.Price).IsRequired()
                        .Property(y => y.Price).ShouldBeGreaterThan(0)
                        .Property(y => y.Unit).IsRequired()
                        .Property(y => y.Unit).ShouldBeGreaterThan(0)
                        .Property(y => y.DepotExclusions).MapFromColumns("Reynolds Dairy", "Waltham Cross")
                ).ToArray();

            var result1 = result.First();
            result1.MeasureId.Should().Be(1);
            result2 = result.Last();
            result2.MeasureId.Should().Be(2);
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
            row1.Row.Should().Be(4);
            var row2 = result.Skip(1).First();
            row2.String.Should().Be("zyx987");
            row2.Int.Should().Be(2);
            row2.DateTime.Should().Be(new DateTime(2015, 10, 9));
            row2.Decimal.Should().Be(9.876);
            row2.Row.Should().Be(5);
            var row3 = result.Skip(2).First();
            row3.String.Should().BeNull();
            row3.Int.Should().BeNull();
            row3.DateTime.Should().BeNull();
            row3.Decimal.Should().Be(5);
            row3.Row.Should().Be(6);
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
                    .UsingHeaderNames(true)
                    .UsingSheet("Prices")
                    .OutputExceptionsTo(exceptionList)
                    .Property(y => y.MinimumOrderQuantity).IsRequired()
                    .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                    .Property(y => y.SupplierDescription).UsesColumnNamed("Supplier Description")
                    .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("dutch description")
                    .Property(y => y.DepotExclusions).MapFromColumns("reynolds dairy", "waltham cross")
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
                    .UsingHeaderNames(true)
                    .UsingSheet("Prices")
                    .Property(y => y.SupplierDescription).UsesColumnNamed("Supplier Description")
                    .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("dutch description")
                    .Property(y => y.DepotExclusions).MapFromColumns("reynolds dairy","waltham cross")
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
        public void NullableBoolCollectionRequiredTest()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MostComplexExample.xlsx");
            var exceptionList = new List<Exception>();
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                    x => x.StartingFromRow(3)
                        .EndingWithRow(3)
                        .UsingHeaderNames(true)
                        .UsingSheet("Prices")
                        .OutputExceptionsTo(exceptionList)
                        .Property(y => y.MinimumOrderQuantity).IsRequired()
                        .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                        .Property(y => y.Vat).IsRequired()
                        .Property(y => y.Vat).ShouldBeOneOf("Standard", "Reduced", "2nd Reduced", "Zero")
                        .Property(y => y.Measure).IsRequired()
                        .Property(y => y.Measure).ShouldBeOneOf("g", "kg", "Each", "lt")
                        .Property(y => y.PslCategory).IsRequired()
                        .Property(y => y.PslCategory).ShouldBeOneOf("Frozen Foods")
                        .Property(y => y.Sku).ShouldBeUnique()
                        .Property(y => y.Sku).IsRequired()
                        .Property(y => y.TranslatedSupplierDescriptions).IsRequired()
                        .Property(y => y.SupplierDescription).UsesColumnNamed("Supplier Description")
                        .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("Dutch Description")
                        .Property(y => y.Price).IsRequired()
                        .Property(y => y.Price).ShouldBeGreaterThan(0)
                        .Property(y => y.Unit).IsRequired()
                        .Property(y => y.Unit).ShouldBeGreaterThan(0)
                        .Property(y => y.DepotExclusions).MapFromColumns("Reynolds Dairy", "Waltham Cross")
                ).ToArray();
            
            result.Count().Should().Be(1);
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
        public void MostComplexExampleWorks()
        {
            var testSpreadsheetLocation = TestHelper.TestsheetPath("MostComplexExample.xlsx");
            var exceptionList = new List<Exception>();
            var result =
                testSpreadsheetLocation.ExcelToEnumerable<ComplexExampleTestClass>(
                    x => x.StartingFromRow(3)
                    .UsingHeaderNames(true)
                    .UsingSheet("Prices")
                    .OutputExceptionsTo(exceptionList)
                    .Property(y => y.MinimumOrderQuantity).IsRequired()
                    .Property(y => y.MinimumOrderQuantity).ShouldBeGreaterThan(0)
                    .Property(y => y.Vat).IsRequired()
                    .Property(y => y.Vat).ShouldBeOneOf("Standard", "Reduced", "2nd Reduced", "Zero")
                    .Property(y => y.Measure).IsRequired()
                    .Property(y => y.Measure).ShouldBeOneOf("g", "kg", "Each", "lt")
                    .Property(y => y.PslCategory).IsRequired()
                    .Property(y => y.PslCategory).ShouldBeOneOf("Frozen Foods")
                    .Property(y => y.Sku).ShouldBeUnique()
                    .Property(y => y.Sku).IsRequired()
                    .Property(y => y.TranslatedSupplierDescriptions).IsRequired()
                    .Property(y => y.SupplierDescription).UsesColumnNamed("Supplier Description")
                    .Property(y => y.TranslatedSupplierDescriptions).MapFromColumns("Dutch Description")
                    .Property(y => y.Price).IsRequired()
                    .Property(y => y.Price).ShouldBeGreaterThan(0)
                    .Property(y => y.Unit).IsRequired()
                    .Property(y => y.Unit).ShouldBeGreaterThan(0)
                    .Property(y => y.DepotExclusions).MapFromColumns("Reynolds Dairy","Waltham Cross")
                ).ToArray();

            //Check that all the minimum order quantities have a value:
            var minimumOrderQuantityWithNoValue = result.FirstOrDefault(x => !x.MinimumOrderQuantity.HasValue);
            minimumOrderQuantityWithNoValue.Should().BeNull();

            //Check that there's a single dupe exception:
            var dupeException = exceptionList.Single(x => x is ExcelToEnumerableSheetException);
            dupeException.Message.Should().Contain("5124EA");

            var cellExceptions = exceptionList.Where(x => x is ExcelToEnumerableCellException).Cast<ExcelToEnumerableCellException>();
            var noMinimumOrderQuantityException = cellExceptions.First(x => x.RowNumber == 7 && x.Column == "H");
            var invalidMinimumOrderQuantityException = cellExceptions.First(x => x.RowNumber == 8 && x.Column == "H");
            var negativeMinimumOrderQuantityException = cellExceptions.First(x => x.RowNumber == 9 && x.Column == "H");
            var missingVatException = cellExceptions.First(x => x.RowNumber == 10 && x.Column == "I");
            var invalidVatValueException = cellExceptions.First(x => x.RowNumber == 11 && x.Column == "I");
            var missingMeasureValue = cellExceptions.First(x => x.RowNumber == 12 && x.Column == "M");
            var invalidMeasureValue = cellExceptions.First(x => x.RowNumber == 13 && x.Column == "M");
            var missingPslCategory = cellExceptions.First(x => x.RowNumber == 14 && x.Column == "C");
            var invalidPslCategory = cellExceptions.First(x => x.RowNumber == 15 && x.Column == "C");
            var missingSku = cellExceptions.First(x => x.RowNumber == 16 && x.Column == "A");
            var missingSupplierDescription = cellExceptions.First(x => x.RowNumber == 17 && x.Column == "E");
            var missingDutchDescription = cellExceptions.First(x => x.RowNumber == 18 && x.Column == "F");
            var missingPrice = cellExceptions.First(x => x.RowNumber == 19 && x.Column == "G");
            var negativePrice = cellExceptions.First(x => x.RowNumber == 20 && x.Column == "G");
            var badCaseQuantity = cellExceptions.First(x => x.RowNumber == 21 && x.Column == "J");
            var badPackQuantity = cellExceptions.First(x => x.RowNumber == 22 && x.Column == "K");
            var missingUnit = cellExceptions.First(x => x.RowNumber == 23 && x.Column == "L");
            var badUnit = cellExceptions.First(x => x.RowNumber == 24 && x.Column == "L");
            var negativeUnit = cellExceptions.First(x => x.RowNumber == 25 && x.Column == "L");
            var badLastPeriodVolume = cellExceptions.First(x => x.RowNumber == 26 && x.Column == "Q");
            var badDepotValue = cellExceptions.First(x => x.RowNumber == 27 && x.Column == "R");

            exceptionList.Count.Should().Be(22);

            result.Count().Should().Be(2);

            var row4966cs = result.First();
            row4966cs.Sku.Should().Be("4966CS");
            row4966cs.SupplierCategory.Should().Be("CONNOISSEURS CHOICE");
            row4966cs.PslCategory.Should().Be("Frozen Foods");
            row4966cs.Store.Should().Be("Chilled");
            row4966cs.TranslatedSupplierDescriptions.Last().Should().Be("X");
            row4966cs.Price.Should().Be(1.01M);
            row4966cs.MinimumOrderQuantity.Should().Be(1);
            row4966cs.Vat.Should().Be("Standard");
            row4966cs.CaseQty.Should().Be(1);
            row4966cs.PackQty.Should().Be(1);
            row4966cs.Unit.Should().Be(1);
            row4966cs.Measure.Should().Be("Each");
            row4966cs.UomDescription.Should().Be("1 X TRAY");
            row4966cs.Origin.Should().BeNull();
            row4966cs.Brand.Should().BeNull();
            row4966cs.LastPeriodVolume.Should().BeNull();
            row4966cs.DepotExclusions.First().Should().BeTrue();
            row4966cs.DepotExclusions.Last().Should().BeFalse();
            
            var row4512cs = result.Last();
            row4512cs.Sku.Should().Be("4512CS");
            row4512cs.SupplierCategory.Should().Be("CONNOISSEURS CHOICE");
            row4512cs.PslCategory.Should().Be("Frozen Foods");
            row4512cs.Store.Should().Be("Chilled");
            row4512cs.TranslatedSupplierDescriptions.Last().Should().Be("X");
            row4512cs.Price.Should().Be(23.59M);
            row4512cs.MinimumOrderQuantity.Should().Be(1);
            row4512cs.Vat.Should().Be("Reduced");
            row4512cs.CaseQty.Should().Be(1);
            row4512cs.PackQty.Should().Be(1);
            row4512cs.Unit.Should().Be(5);
            row4512cs.Measure.Should().Be("kg");
            row4512cs.UomDescription.Should().Be("1X 5KG");
            row4512cs.Origin.Should().BeNull();
            row4512cs.Brand.Should().BeNull();
            row4512cs.LastPeriodVolume.Should().BeNull();
            row4512cs.DepotExclusions.Should().BeNull();
        }
        
        /// <summary>
        /// CSH 11012019 This is not a test, it's just somewhere to check that the code in the documentation example actually compiles
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
    }

}