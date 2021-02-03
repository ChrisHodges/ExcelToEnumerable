namespace ExcelToEnumerable.Tests.TestClasses
{
    public class EnumsTestClass
    {
        public enum ParseFromStringsEnum
        {
            Value1 = 1,
            Value2 = 2
        }

        public enum ParseFromIntsEnum
        {
            Ten = 10,
            Twenty = 20
        }

        public ParseFromStringsEnum ParseFromStrings { get; set; }
        public ParseFromIntsEnum? ParseFromInts { get; set; }
    }
    
    public class StringsEnumsTestClass
    {
        public enum ParseFromStringsEnum
        {
            Value1 = 1,
            Value2 = 2
        }

        public ParseFromStringsEnum ParseFromStrings { get; set; }
    }
}