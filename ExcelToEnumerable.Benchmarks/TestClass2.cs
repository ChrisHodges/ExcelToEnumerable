using ExcelDataReader;
using FileHelpers;

namespace ExcelToEnumerable.Benchmarks
{
    [DelimitedRecord("|")]
    public class TestClass2
    {
        public TestClass2()
        {
        }

        public TestClass2(IExcelDataReader reader)
        {
            Sku = reader.IsDBNull(0) ? default : reader.GetString(0);
            SupplierCategory = reader.IsDBNull(6) ? default : reader.GetString(1);
            Store = reader.IsDBNull(2) ? default : reader.GetString(2);
            SupplierDescription = reader.IsDBNull(3) ? default : reader.GetString(3);
            Price = reader.GetDouble(4);
            VAT = reader.IsDBNull(5) ? default : reader.GetString(5);
            CaseQty = reader.IsDBNull(6) ? default(double?) : reader.GetDouble(6);
            PackQty = reader.IsDBNull(7) ? default(double?) : reader.GetDouble(7);
            Unit = reader.GetDouble(8);
            Measure = reader.IsDBNull(9) ? default : reader.GetString(9);
            UomDescription = reader.IsDBNull(10) ? default : reader.GetString(10);
            Origin = reader.IsDBNull(11) ? default : reader.GetString(11);
            Brand = reader.IsDBNull(12) ? default : reader.GetString(12);
            LastPeriodVolume = reader.IsDBNull(13) ? default(double?) : reader.GetDouble(13);
        }

        public string Sku { get; set; }
        public string SupplierCategory { get; set; }
        public string Store { get; set; }
        public string SupplierDescription { get; set; }
        public double Price { get; set; }
        public string VAT { get; set; }
        public double? CaseQty { get; set; }
        public double? PackQty { get; set; }
        public double Unit { get; set; }
        public string Measure { get; set; }
        public string UomDescription { get; set; }
        public string Origin { get; set; }
        public string Brand { get; set; }
        public double? LastPeriodVolume { get; set; }
    }
}