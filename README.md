# ExcelToEnumerable
[![NuGet](https://img.shields.io/nuget/v/ExcelToEnumerable)](https://www.nuget.org/packages/ExcelToEnumerable)

A high-performance library for mapping Excel tables to a list of objects.

Please fork and submit pull requests to the `development` branch.

### Installation
Install via Nuget, either the Package Manager Console:

```
PM> Install-Package ExcelToEnumerable
```
or the .NET CLI:

```
> dotnet add package ExcelToEnumerable
```

### Quick Start
To get an enumerable of type T from a .xlsx file use the `ExcelToEnumerable<T>` extension method on a file path, or a stream:

```C#
IEnumerable<T> myList = "path/to/ExcelFile.xlsx".ExcelToEnumerable<T>();

IEnumerable<T> listFromStream = myStream.ExcelToEnumerable<T>();
```

### A Simple Example
If you've got an .xlsx spreadsheet that looks like this:

|   |  A               | B           | C        | D            |
|---|------------------|-------------|----------|--------------|
| 1 | **Name**       | **Colour**  | **Legs** | **TopSpeed** |
| 2 | Horse            | Brown       | 4        | 70.76        |
| 3 | African Wild Dog | Black       | 4        | 71           |
| 4 | House Spider     | Black       | 8        | 1.9          |
| 5 | Peregrine Falcon | Grey        | 0        | 242          |

you can use **ExcelToEnumerable** to map it to an `Animal` class like this:

```C#
public class Animal
{
  public string Name { get; set; }
  public string Colour { get; set; }
  public int Legs { get; set; }
  public decimal TopSpeed { get; set; }
}
```
by using `ExcelToEnumerable<Animal>` extension method with zero configuration:

```C#
IEnumerable<Animal> animals = "path/to/ExcelFile.xlsx".ExcelToEnumerable<Animal>();
```
By default, **ExcelToEnumerable** assumes that:

* The data is on the first sheet in the workbook.
* There is a header on row 1 that matches the properties of the class you're mapping to. (The order of the columns doesn't matter
* The data begins on row 2.

### Configuration
**ExcelToEnumerable** has a fluent options interface and is highly configurable:

```C#
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
```

### Performance
We believe **ExcelToEnumerable** is the fastest way to map Excel data to a collection of objects in .Net. We've included a [Benchmark Dot Net](https://github.com/dotnet/BenchmarkDotNet) project in the repo to compare **ExcelToEnumerable** performance against other libraries. For comparison, here are the results from running the benchmark on a 2.7 GHz Quad-Core Intel Core i7 laptop in Windows 10:

| Package             | Operation |
|---------------------|----------:|
| *ExcelToEnumerable* | *30.86 ms*|
| ExcelEntityMapper   | 176.03 ms |
| ExcelNPOIStorage    | 212.79 ms |
| ExcelDataReader     | 239.45 ms |

### Support
**ExcelToEnumerable** is currently in beta. If you discover a bug please [raise an issue](https://github.com/ChrisHodges/ExcelToEnumerable/issues/new) or better still [fork the project](https://help.github.com/en/github/getting-started-with-github/fork-a-repo) and raise a Pull Request if you can fix the issue yourself.

