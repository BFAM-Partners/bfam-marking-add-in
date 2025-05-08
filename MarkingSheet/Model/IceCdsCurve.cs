using System;

namespace MarkingSheet.Model
{
    internal class IceCdsCurve
    {
        public string Name { get; set; }
        public string Ticker { get; set; }
        public int SophisCurveSicovam { get; set; }
        public string Seniority { get; set; }
        public string DocClause { get; set; }
        public string Currency { get; set; }
        public double? OneYear { get; set; }
        public double? ThreeYear { get; set; }
        public double? FiveYear { get; set; }
        public double? SevenYear { get; set; }
        public double? TenYear { get; set; }
        public bool isIndex { get; set; }
        public DateTime? IceCurveDate { get; set; }
    }
}
