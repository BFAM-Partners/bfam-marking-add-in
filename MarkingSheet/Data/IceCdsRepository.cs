using MarkingSheet.Model;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MarkingSheet.Data
{
    class IceCdsRepository
    {
        private static string _kartConnectionString = "Host=kart-db.prd.bfam-partners.com;Username=toby;Password=4fFhN7CqMuQnoyumylXh;Database=postgres";

        private static string _eqlinkedConnectionString = "Host=eq-linked-shared.prd.bfam-partners.com;Username=credit_api_user;Password='>7,V1;Ag2Q)w';Database=credit_api";

        private static Dictionary<string, string> RestructuringCodes = new Dictionary<string, string>()
        {
            {"XXXXXX","CR"},
            {"FR","CR14"},
            {"MHO","MH14"},
            {"MMR","MM14"},
            {"MR14","MR14"},
            {"NR","XR14"},
        };

        private static Dictionary<string, string> Seniority = new Dictionary<string, string>()
        {
            {"Secured Debt","Secured"},
            {"Senior Unsecured", "SeniorUnsecured"},
            {"Subordinated or Senior Sub","Subordinated"},
            {"Junior Subordinated","JuniorSubordinated"},
            {"XXXXXX","SeniorLossAbsorbingCapacity"},
        };

        public static void EnrichMarkingSheetCdsCurveWithIceCds(IEnumerable<MarkingSheetCds> markingSheetCurves)
        {
            var serverString = _eqlinkedConnectionString;            

            var eqlDataSource = NpgsqlDataSource.Create(serverString);
            NpgsqlConnection eqlConnection = eqlDataSource.OpenConnection();
            var query = "SELECT * FROM credit_api.ice_cds_curve AS c LEFT JOIN credit_api.ice_cds_curve_point AS p ON c.id = p.cds_curve_id " +
                        "WHERE c.id in ( " +
                            "SELECT c1.id " +
                            "FROM credit_api.ice_cds_curve c1 LEFT JOIN credit_api.ice_cds_curve c2 ON ( " +
                            "c1.ticker = c2.ticker " +
                            "AND c1.seniority = c2.seniority " +
                            "AND (c1.doc_clause = c2.doc_clause OR :is_index = true)" +
                            "AND c1.currency = c2.currency " +
                            "AND c1.file_modified_time < c2.file_modified_time " +
                            ") " +
                            "WHERE c2.file_modified_time IS NULL) " +
                        "AND c.ticker = :ticker " +
                        "AND c.seniority = :seniority " +
                        "AND (c.doc_clause = :doc_clause OR :is_index = true) " +
                        "AND c.currency = :currency ";

            Parallel.ForEach(markingSheetCurves, new ParallelOptions { MaxDegreeOfParallelism = 5 }, curve =>
            {
                if(curve.Seniority == null)
                {
                    throw new Exception($"Error for Ticker: {curve.Ticker}. Position (Sicovam: {curve.SwapSicovam}) is missing 'Seniority' in Sophis. Please add Seniority in Sophis and then refresh this sheet after a few minutes.");
                }

                var iceCurve = new IceCdsCurve
                {
                    SophisCurveSicovam = curve.CurveSicovam,
                    DocClause = curve.DocClause,
                    Seniority = curve.Seniority,
                    Currency = curve.Currency,
                    Ticker = curve.Ticker,
                    isIndex = curve.isIndex
                };
                curve.iceCdsCurve = iceCurve;

                var curveQuery = eqlDataSource.CreateCommand(query);

                if (iceCurve.Ticker == null || iceCurve.Currency == null ||
                    (!iceCurve.isIndex && iceCurve.DocClause == null))
                {
                    return;
                }

                curveQuery.Parameters.AddWithValue("currency", iceCurve.Currency);
                curveQuery.Parameters.AddWithValue("ticker", iceCurve.Ticker);
                curveQuery.Parameters.AddWithValue("doc_clause", iceCurve.DocClause != null ? (RestructuringCodes.TryGetValue(iceCurve.DocClause, out string doc_clause) ? doc_clause : "") : "");
                curveQuery.Parameters.AddWithValue("is_index", iceCurve.isIndex);
                curveQuery.Parameters.AddWithValue("seniority", Seniority.TryGetValue(iceCurve.Seniority, out string seniority) ? seniority : "");

                var reader = curveQuery.ExecuteReader();
                {
                    while (reader.Read())
                    {
                        var tenor = reader.GetString(reader.GetOrdinal("tenor"));
                        var curveDate = reader.GetDateTime(reader.GetOrdinal("file_modified_time"));
                        TimeZoneInfo hkZone = TimeZoneInfo.FindSystemTimeZoneById("China Standard Time");
                        DateTime hktCurveTimestamp = TimeZoneInfo.ConvertTimeFromUtc(curveDate, hkZone);
                        var name = reader.GetString(reader.GetOrdinal("name"));

                        if (tenor == "1Y")
                        {
                            iceCurve.OneYear = reader.GetDouble(reader.GetOrdinal("conventional_spread"));
                        }
                        else if (tenor == "3Y")
                        {
                            iceCurve.ThreeYear = reader.GetDouble(reader.GetOrdinal("conventional_spread"));
                        }
                        else if (tenor == "5Y")
                        {
                            iceCurve.FiveYear = reader.GetDouble(reader.GetOrdinal("conventional_spread"));
                        }
                        else if (tenor == "7Y")
                        {
                            iceCurve.SevenYear = reader.GetDouble(reader.GetOrdinal("conventional_spread"));
                        }
                        else if (tenor == "10Y")
                        {
                            iceCurve.TenYear = reader.GetDouble(reader.GetOrdinal("conventional_spread"));
                        }
                        iceCurve.IceCurveDate = hktCurveTimestamp;
                        iceCurve.Name = name;
                    }
                }                
            });
            eqlConnection.Dispose();
            eqlDataSource.Dispose();
        }

    }

}
