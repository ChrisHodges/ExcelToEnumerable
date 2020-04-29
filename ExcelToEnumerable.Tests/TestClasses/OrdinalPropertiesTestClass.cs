namespace ExcelToEnumerable.Tests
{
    public class OrdinalPropertiesTestClass
    {
        /// <summary>
        /// CSH 28041890 Deliberately putting the columns here in a different order to how the appear on the spreadsheet
        /// to properly test the property=>UseColumnNumber option
        /// </summary>
        public string ColumnC { get; set; }
        
        public string ColumnA { get; set; }
        
        public string ColumnB { get; set; }
        
        public string ColumnAA { get; set; }
        
        public string IgnoreThisProperty { get; set; }
        
        public int Row { get; set; }
    }
}