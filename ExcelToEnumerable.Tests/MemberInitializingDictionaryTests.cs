using System.Collections.Generic;
using Xunit;

namespace ExcelToEnumerable.Tests
{
    public class MemberInitializingDictionaryTests
    {
        [Fact]
        public void AddIndexerWorks()
        {
            // ReSharper disable once CollectionNeverUpdated.Local
            var subject = new MemberInitializingDictionary<string,List<ExcelCellValidator>>();
            subject["KEY"].Add(new ExcelCellValidator());
        }
    }
}