using System;
using ExcelToEnumerable.Attributes;

namespace ExcelToEnumerable.Tests
{
    [StartingFromRow(3)]
    [WithBlankRowBehaviour(BlankRowBehaviour.StopReading)]
    public class BlankRowBehaviourAttributeTest
    {
        public string Sku { get; set; }
        public string SupplierCategory { get; set; }
        public string Store { get; set; }
        public string SupplierDescription { get; set; }
        public double Price { get; set; }
        public string Vat { get; set; }
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