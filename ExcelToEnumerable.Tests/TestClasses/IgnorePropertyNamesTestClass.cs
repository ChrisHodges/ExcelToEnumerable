namespace ExcelToEnumerable.Tests
{
    public class IgnorePropertyNamesTestClass
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }
        
        public string NotOnSpreadsheet { get; set; }
    }
}