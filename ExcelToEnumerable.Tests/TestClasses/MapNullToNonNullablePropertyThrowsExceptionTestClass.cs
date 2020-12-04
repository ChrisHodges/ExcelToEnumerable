namespace ExcelToEnumerable.Tests.TestClasses
{
    public class MapNullToNonNullablePropertyThrowsExceptionTestClass
    {
        public int NotNullable { get; set; }
        public int? Nullable { get; set; }
        public string String { get; set; }
    }
}