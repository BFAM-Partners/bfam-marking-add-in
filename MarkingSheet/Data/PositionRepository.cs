using MarkingSheet.Model;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Amazon.S3;
using Amazon.S3.Model;
using System.IO;
using Amazon.Runtime;
using Newtonsoft.Json.Linq;
using System.Linq;
using Amazon.Runtime.CredentialManagement;
using Amazon;

namespace MarkingSheet
{
    internal class PositionRepository
    {

        // kart-scratch.cdlnjvn2zuyi.ap-east-1.rds.amazonaws.com
        private readonly AmazonS3Client _s3Client;

        public PositionRepository()
        {
            var chain = new CredentialProfileStoreChain();
            AWSCredentials awsCredentials;

            if (chain.TryGetAWSCredentials("creditmarking", out awsCredentials))
            {
                _s3Client = new AmazonS3Client(awsCredentials, RegionEndpoint.APEast1);
            }
            else
            {
                throw new Exception($"Failed to find AWS credentials for profile: 'creditmarking'");
            }
        }

        private static readonly string _kartConnectionString =
    "Host=kart-db.prd.bfam-partners.com;Username=toby;Password=4fFhN7CqMuQnoyumylXh;Database=postgres";
        public JArray ReadS3FileContents(string bucketName, string keyName)
        {
            try
            {
                GetObjectRequest request = new GetObjectRequest
                {
                    BucketName = bucketName,
                    Key = keyName
                };

                using (GetObjectResponse response = _s3Client.GetObject(request))
                using (Stream responseStream = response.ResponseStream)
                using (StreamReader reader = new StreamReader(responseStream))
                {
                    string contents = reader.ReadToEnd();
                    JArray jsonContents = JArray.Parse(contents);
                    return jsonContents;
                }
            }
            catch (AmazonServiceException e)
            {
                Console.WriteLine($"Request ID: {e.RequestId}");
                Console.WriteLine($"Error Code: {e.ErrorCode}");
                Console.WriteLine($"Error Message: {e.Message}");
                Console.WriteLine($"HTTP Status Code: {e.StatusCode}");
                throw;
            }
            catch (Exception e)
            {
                Console.WriteLine($"General Error: {e.Message}");
                throw;
            }
        }

        public static List<Bond> FetchBonds() 
        {

            var serverString = _kartConnectionString;            

            var kartDataSource = NpgsqlDataSource.Create(serverString);
            NpgsqlConnection kartConnection = kartDataSource.OpenConnection();

            var query = "SELECT reference, name, instruments.sicovam, SUM(quantity) as qty "+
                "FROM active_positions " +
                "LEFT JOIN instruments ON active_positions.sicovam = instruments.sicovam " +
                "WHERE instruments.instrument_type = 'bond' GROUP BY instruments.reference, instruments.name, instruments.sicovam";


            var instrumentQuery = kartDataSource.CreateCommand(query);

            var instruments = new List<Bond>();
            var reader = instrumentQuery.ExecuteReader();
            {
                while (reader.Read())
                {
                    var instrument = new Bond
                    {
                        Sicovam = reader.GetInt32(reader.GetOrdinal("sicovam")),
                        Name = reader.GetString(reader.GetOrdinal("name")),
                        Quantity = reader.GetDouble(reader.GetOrdinal("qty")),
                        Isin = reader.GetString(reader.GetOrdinal("reference"))
                    };                    
                    instruments.Add(instrument);
                }
            }

            kartConnection.Dispose();            

            kartDataSource.Dispose();

            return instruments;

        }

        public static List<SophisCdsCurve> FetchSophisCdsCurves()
        {
            var serverString = _kartConnectionString;

            var kartDataSource = NpgsqlDataSource.Create(serverString);
            NpgsqlConnection kartConnection = kartDataSource.OpenConnection();

            var query = "SELECT sicovam, fpml FROM fpmls WHERE sicovam IN (SELECT DISTINCT(instruments.sicovam) " +
                        "FROM active_positions\r\nLEFT JOIN instruments ON active_positions.sicovam = instruments.sicovam " +
                        "WHERE instruments.instrument_type = 'swap' AND instruments.allotment = 'CDS' AND active_positions.quantity != 0)";

            var instrumentQuery = kartDataSource.CreateCommand(query);

            var instruments = new List<SophisCdsCurve>();
            var reader = instrumentQuery.ExecuteReader();
            {
                while (reader.Read())
                {
                    var fpml = reader.GetString(reader.GetOrdinal("fpml"));
                    var cdsSwapSicovam = reader.GetInt32(reader.GetOrdinal("sicovam"));

                    XDocument xdoc = XDocument.Parse(fpml);
                    XmlNamespaceManager xnm = new XmlNamespaceManager(new NameTable());
                    xnm.AddNamespace("ns0", "http://www.sophis.net/Instrument");
                    xnm.AddNamespace("ns1", "http://sophis.net/sophis/common");

                    var curveSicovam = xdoc.XPathSelectElement("./ns0:swap/ns0:issuer/ns0:sophis", xnm)?.Value;

                    var curveReference = xdoc.XPathSelectElement("./ns0:swap/ns0:issuer/ns0:reference[@ns0:name = 'Reference']", xnm)?.Value;

                    var senority = xdoc.XPathSelectElement("./ns0:swap/ns0:receivingLeg/ns0:creditLeg/ns0:obligations/ns0:seniority", xnm)?.Value;
                    
                    var restructuringType = xdoc.XPathSelectElement("./ns0:swap/ns0:identifier/ns0:reference[@ns0:name = 'RestructuringType']", xnm)?.Value;

                    var redCode = xdoc.XPathSelectElement("./ns0:swap/ns0:identifier/ns0:reference[@ns0:name = 'Red']", xnm)?.Value;

                    var currency = xdoc.XPathSelectElement("./ns0:swap/ns0:currency", xnm)?.Value;
                    
                    var productType = xdoc.XPathSelectElement("./ns0:swap/ns0:productType", xnm)?.Value;
                    var isIndex = productType.Contains("Basket");

                    var instrument = new SophisCdsCurve
                    {
                        CurveSicovam = Convert.ToInt32(curveSicovam),
                        Ticker = curveReference,
                        SwapSicovam = cdsSwapSicovam,
                        DocClause = restructuringType,
                        Seniority = senority,
                        Currency = currency,
                        isIndex = isIndex
                    };
                    
                    instruments.Add(instrument);
                }
            }

            kartConnection.Dispose();

            kartDataSource.Dispose();

            return instruments;

        }
        public string GetLatestFileName(string bucketName, string filePattern)
        {
            try
            {
                // Get today's date to start the search
                DateTime today = DateTime.UtcNow.Date;
                bool fileFound = false;
                string latestFileName = null;

                // Keep checking previous dates until the file is found
                while (!fileFound)
                {
                    string prefix = today.ToString("yyyy-MM-dd");

                    // Define the request to list objects in the bucket for the specific date
                    ListObjectsV2Request request = new ListObjectsV2Request
                    {
                        BucketName = bucketName,
                        Prefix = prefix
                    };

                    // Execute the request synchronously
                    ListObjectsV2Response response = _s3Client.ListObjectsV2Async(request).Result;

                    // Filter the objects based on the file pattern
                    foreach (S3Object entry in response.S3Objects)
                    {
                        if (entry.Key.EndsWith(filePattern))
                        {
                            latestFileName = entry.Key;
                            fileFound = true;
                            break;
                        }
                    }

                    // If the file is not found, go to the previous day
                    if (!fileFound)
                    {
                        today = today.AddDays(-1);
                    }
                }

                return latestFileName;
            }
            catch (AmazonServiceException e)
            {
                Console.WriteLine($"Error encountered on server. Message:'{e.Message}' when listing objects");
                throw;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unknown encountered on server. Message:'{e.Message}' when listing objects");
                throw;
            }
        }
        public static List<ConvertibleBond> FetchConvertibles()
        {

            // 1. Get data from Kart
            var serverString = _kartConnectionString;

            string convertibleBondsQuery = @"
                SELECT DISTINCT
                    instruments.reference,
                    instruments.underlying_reference,
                    folio.name as portfolio_name,
                    instruments.name,
                    instruments.sicovam,
                    instruments.notional_amount,
                    active_positions.number_of_securities,                    
                    active_positions.position * active_positions.fx_rate as position_usd,
                    active_positions.currency,
                    instruments.underlying_sicovam,
                    active_positions.gog_underlying_spot_price as underlying_spot,
                    active_positions.theoretical,
                    active_positions.last,
                    historical_active_positions.theoretical as historical_theoretical,
                    historical_active_positions.last as historical_last,
                    active_positions.delta,
                    active_positions.gog_credit_spread,
                    repo_rate.repo_rate,
                    active_positions.volatility,
                    'convertibleBond' as instrument_type
                FROM instruments
                LEFT JOIN active_positions ON active_positions.sicovam = instruments.sicovam
                LEFT JOIN historical_active_positions ON active_positions.position_id = historical_active_positions.position_id
                LEFT JOIN (
                        SELECT id, string_agg(name, ', ') as name
                        FROM portfolios
                        GROUP BY id
                    ) folio ON folio.id = active_positions.folio_id
                LEFT JOIN (
                    SELECT sicovam, repo_rate
                    FROM repo_rate
                    WHERE (sicovam, maturity) IN (
                        SELECT sicovam, MAX(maturity)
                        FROM repo_rate
                        GROUP BY sicovam
                    )
                ) repo_rate ON repo_rate.sicovam = instruments.underlying_sicovam
                WHERE instruments.instrument_type = 'convertibleBond'
                    AND instruments.underlying_reference not ilike '% CH Equity%'
                    AND active_positions.quantity <> 0
                    AND DATE(historical_active_positions.date) = (
                        SELECT CASE
                            WHEN EXTRACT(DOW FROM CURRENT_DATE) = 1 THEN CURRENT_DATE - INTERVAL '3 days'
                            WHEN EXTRACT(DOW FROM CURRENT_DATE) = 0 THEN CURRENT_DATE - INTERVAL '2 days'
                            ELSE CURRENT_DATE - INTERVAL '1 day'
                        END 
                    )
            ";

            string ascotsQuery = @"
                SELECT DISTINCT
                    instruments.underlying_reference as reference,
                    underlying.underlying_reference as underlying_reference,
                    folio.name as portfolio_name,
                    instruments.name,
                    instruments.sicovam,
                    instruments.notional_amount,
                    active_positions.number_of_securities,
                    active_positions.position * active_positions.fx_rate as position_usd,
                    active_positions.currency,
                    instruments.underlying_sicovam,
                    active_positions.gog_underlying_spot as underlying_spot,
                    active_positions.gog_repo_margin as repo_rate,
                    active_positions.gog_volatility as volatility,
                    active_positions.gog_credit_spread,
                    active_positions.gog_cb_price as theoretical,
                    active_positions.last,
                    historical_active_positions.gog_cb_price as historical_theoretical,
                    historical_active_positions.last as historical_last,
                    active_positions.delta,
                    instruments.instrument_type
                FROM instruments
                LEFT JOIN active_positions ON active_positions.sicovam = instruments.sicovam
                LEFT JOIN historical_active_positions ON active_positions.position_id = historical_active_positions.position_id
                LEFT JOIN (
                        SELECT id, string_agg(name, ', ') as name
                        FROM portfolios
                        GROUP BY id
                    ) folio ON folio.id = active_positions.folio_id
                LEFT JOIN instruments as underlying ON instruments.underlying_sicovam = underlying.sicovam
                WHERE (instruments.instrument_type = 'ascot' or (instruments.instrument_type = 'swap' and instruments.name ilike 'ES_CBSW%'))
                    AND underlying.underlying_reference not ilike '% CH Equity%'
                    AND active_positions.quantity <> 0
                    AND DATE(historical_active_positions.date) =  (
                        SELECT CASE
                            WHEN EXTRACT(DOW FROM CURRENT_DATE) = 1 THEN CURRENT_DATE - INTERVAL '3 days'
                            WHEN EXTRACT(DOW FROM CURRENT_DATE) = 0 THEN CURRENT_DATE - INTERVAL '2 days'
                            ELSE CURRENT_DATE - INTERVAL '1 day'
                        END 
                    )
            ";

            string[] queries = { convertibleBondsQuery, ascotsQuery };

            var instruments = new List<ConvertibleBond>();

            using (var conn = new NpgsqlConnection(serverString))
            {
                conn.Open();
                foreach (string query in queries)
                {
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var instrument = new ConvertibleBond
                                {
                                    Isin = reader["reference"] as string,
                                    Name = reader["name"] as string,
                                    Portfolio = reader["portfolio_name"] as string,
                                    InstrumentType = reader["instrument_type"] as string,
                                    UnderlyingReference = reader["underlying_reference"] as string,
                                    Sicovam = reader.IsDBNull(reader.GetOrdinal("sicovam")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("sicovam")),
                                    Volatility = reader.IsDBNull(reader.GetOrdinal("volatility")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("volatility")),
                                    Spread = reader.IsDBNull(reader.GetOrdinal("gog_credit_spread")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("gog_credit_spread")) * 100,
                                    Theoretical = reader.IsDBNull(reader.GetOrdinal("theoretical")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("theoretical")),
                                    Last = reader.IsDBNull(reader.GetOrdinal("last")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("last")),
                                    HistoricalTheoretical = reader.IsDBNull(reader.GetOrdinal("historical_theoretical")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("historical_theoretical")),
                                    HistoricalLast = reader.IsDBNull(reader.GetOrdinal("historical_last")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("historical_last")),
                                    UnderlyingPrice = reader.IsDBNull(reader.GetOrdinal("underlying_spot")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("underlying_spot")),
                                    Delta = reader.IsDBNull(reader.GetOrdinal("delta")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("delta")),
                                    Borrow = reader.IsDBNull(reader.GetOrdinal("repo_rate")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("repo_rate")),
                                    Currency = reader["currency"] as string,
                                    PositionUSD = reader.IsDBNull(reader.GetOrdinal("position_usd")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("position_usd")),
                                };
                                instruments.Add(instrument);
                            }
                        }
                    }
                }
            }

            // 2. Get data from S3
            string markBucket = "bfam-prd-eq-linked-parsed-cb-quotes";
            PositionRepository fetchPositionsInstance = new PositionRepository();

            string bestQuoteFile = fetchPositionsInstance.GetLatestFileName(markBucket, "compositeQuotes.json");
            JArray bestQuotes = fetchPositionsInstance.ReadS3FileContents(markBucket, bestQuoteFile);

            Dictionary<string, JToken> bestQuotesDict = bestQuotes.ToDictionary(bq => (string)bq["isinOrFigi"], bq => bq);

            foreach (var instrument in instruments)
            {
               
                if (bestQuotesDict.TryGetValue(instrument.Isin, out var matchingQuote))
                {
                    JToken bestAsk = matchingQuote["bestAsk"];
                    JToken bestBid = matchingQuote["bestBid"];

                    string bestBidIssuer = (string)bestBid["source"];
                    double? bestBidBid = (double?)bestBid["bid"];
                    double? bestBidRef = (double?)bestBid["refPrice"];

                    string bestAskIssuer = (string)bestAsk["source"];
                    double? bestAskAsk = (double?)bestAsk["ask"];
                    double? bestAskRef = (double?)bestAsk["refPrice"];

                    instrument.BestBidIssuer = bestBidIssuer;
                    instrument.BestBidBid = bestBidBid;
                    instrument.BestBidRef = bestBidRef;
                    instrument.BestAskIssuer = bestAskIssuer;
                    instrument.BestAskAsk = bestAskAsk;
                    instrument.BestAskRef = bestAskRef;

                    // nuke the Best Ask and Best Bid
                    double? bestBidNuked = instrument.BestBidBid + (instrument.UnderlyingPrice - instrument.BestBidRef) * instrument.Delta;
                    double? bestAskNuked = instrument.BestAskAsk + (instrument.UnderlyingPrice - instrument.BestAskRef) * instrument.Delta;
                    instrument.BestBidNuked = bestBidNuked;
                    instrument.BestAskNuked = bestAskNuked;
                }
            }

            return instruments;

        }

    }
}
