namespace MarkingSheet
{
    internal class ConvertibleBond
    {
        public string Isin { get; set; }

        public string UnderlyingReference { get; set; }

        public int? Sicovam { get; set; }

        public string Name { get; set; }

        public string Portfolio { get; set; }

        public string InstrumentType { get; set; }

        public string BestBidIssuer { get; set; }

        public double? BestBidBid { get; set; }

        public double? BestBidNuked { get; set; }

        public double? BestAskNuked { get; set; }

        public double? BestBidRef { get; set; }

        public double? BestAskRef { get; set; }

        public string BestAskIssuer { get; set; }

        public double? BestAskAsk { get; set; }

        public double? Volatility { get; set; }

        public double? Spread { get; set; }

        public double? Theoretical { get; set; }
        
        public double? Last { get; set; }

        public double? HistoricalTheoretical { get; set; }
        
        public double? HistoricalLast { get; set; }

        public double? UnderlyingPrice { get; set; }

        public double? Delta { get; set; }

        public double? Borrow { get; set; }

        public double? PositionUSD { get; set; }

        public string Currency { get; set; }
        
        public double? BondPosition { get; set; }

        public double? AscotPosition { get; set; }

        public override bool Equals(object obj)
        {
            var other = obj as ConvertibleBond;
            if (other == null)
                return false;
            return this.Isin == other.Isin;
        }

        public override int GetHashCode()
        {
            return Isin != null ? Isin.GetHashCode() : 0;
        }

    }
}