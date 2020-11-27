namespace ExcelToEnumerable.Tests.TestClasses
{
    public class AllPropertiesMustBeMappedToColumnsTestClass
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }

        public string NotOnSpreadsheet { get; set; }
    }
}