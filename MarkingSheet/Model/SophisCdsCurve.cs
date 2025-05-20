using System;

namespace MarkingSheet.Model
{
    internal class SophisCdsCurve
    {
        public string Ticker { get; set; }
        public string Seniority { get; set; }
        public int CurveSicovam { get; set; }
        public string DocClause { get; set; }
        public string Currency { get; set; }
        public int SwapSicovam { get; set; }
        public bool isIndex { get; set; }        

        public override bool Equals(object obj)
        {
            var other = obj as SophisCdsCurve;
            if (other == null)
                return false;

            return this.Ticker == other.Ticker &&
                   this.Seniority == other.Seniority &&
                   this.CurveSicovam == other.CurveSicovam;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Ticker, Seniority, CurveSicovam);
        }

    }
}
