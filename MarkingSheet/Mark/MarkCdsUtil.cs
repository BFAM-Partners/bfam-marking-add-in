using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Amazon;
using Amazon.Lambda;
using Amazon.Lambda.Model;
using Amazon.Runtime.CredentialManagement;
using MarkingSheet.Model;

namespace MarkingSheet.Mark
{
    internal class MarkCdsUtil
    {        
        public static string SubmitMarks(CdsMarkParameters markedCurve)
        {
            // 
            var chain = new CredentialProfileStoreChain();
            chain.TryGetAWSCredentials("creditmarking", out var credentials);
            var lambdaConfig = new AmazonLambdaConfig() { RegionEndpoint = RegionEndpoint.APEast1 };
            var lambdaClient = new AmazonLambdaClient(credentials, lambdaConfig);

            CdsMarkPayload value = new CdsMarkPayload() {
                Parameters = markedCurve
            };
            var payload = JsonSerializer.Serialize(value);

            var lambdaRequest = new InvokeRequest
            {
                FunctionName = "voltool-vt-lambda-generic",
                Payload = payload
            };

            var response = lambdaClient.Invoke(lambdaRequest);
            if (response != null)
            {
                using (var sr = new StreamReader(response.Payload))
                {
                    string responseText = sr.ReadToEnd().Trim();
                    if (responseText == "200")
                    {
                        return "OK";
                    }
                    else
                    {
                        return responseText;
                    }
                }
            }
            return $"No response from CDS marking service for CDS {markedCurve.Reference}";
        }

        class CdsMarkPayload
        {
            public string Function { get; } = "SophisCDSCurveUpdate";
            public CdsMarkParameters Parameters { get; set; } = new CdsMarkParameters();
        }

        internal class CdsMarkParameters
        {
            public string Reference { get; set; }
            public List<CdsPoint> Points { get; set; }
        }

        public class CdsPoint
        {
            public string Seniority { get; set; }
            public double PeriodMultiplier { get; set; }
            public string PeriodEnum { get; set; } = "Year";
            public double? Rate { get; set; }
        }


    }
}
