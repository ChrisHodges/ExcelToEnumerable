using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests.TestClasses
{
    // ReSharper disable once ClassNeverInstantiated.Global
    public class RequiredColumnsAttributeTestClass
    {
        public string OptionalProperty { get; set; }
        [RequiredColumn]
        public string RequiredProperty { get; set; }
    }
}