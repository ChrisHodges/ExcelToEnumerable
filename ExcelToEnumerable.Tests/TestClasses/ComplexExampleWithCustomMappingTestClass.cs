using System.Collections.Generic;

namespace ExcelToEnumerable.Tests
{
    public class ComplexExampleWithCustomMappingTestClass
    {
        public string Sku { get; set; }
        public string SupplierCategory { get; set; }
        public string PslCategory { get; set; }
        public string Store { get; set; }
        public List<string> TranslatedSupplierDescriptions { get; set; }
        public decimal Price { get; set; }
        public int? MinimumOrderQuantity { get; set; }
        public string Vat { get; set; }
        public int? CaseQty { get; set; }
        public decimal? PackQty { get; set; }
        public decimal? Unit { get; set; }
        public int MeasureId { get; set; }
        public string UomDescription { get; set; }
        public string Origin { get; set; }
        public string Brand { get; set; }
        public decimal? LastPeriodVolume { get; set; }
        public List<bool?> DepotExclusions { get; set; }
    }
}