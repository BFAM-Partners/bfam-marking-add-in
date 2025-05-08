using Amazon;
using Amazon.Lambda;
using Amazon.Lambda.Model;
using Amazon.Runtime;
using Amazon.Runtime.CredentialManagement;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace MarkingSheet
{
    internal class MarkBondUtil
    {        
        public static void SubmitMarks(IEnumerable<Mark> marks)
        {
            // 
            var chain = new CredentialProfileStoreChain();
            chain.TryGetAWSCredentials("creditmarking", out var credentials);
            var lambdaConfig = new AmazonLambdaConfig() { RegionEndpoint = RegionEndpoint.APEast1 };
            var lambdaClient = new AmazonLambdaClient(credentials, lambdaConfig);

            MarkPayload value = new MarkPayload()
            {
                Parameters = new MarkParameters()
                {
                    list = marks
                }
            };
            var payload = JsonSerializer.Serialize(value);

            var lambdaRequest = new InvokeRequest
            {
                /*
                 { "Function":"SophisUpdate","Parameters":{"list":[{"sicovam":71696406,"fieldName":"Last","dateTime":"2024-01-29","level":190.41}]}} 
                 */ 
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
                        MessageBox.Show("Marks posted for processing");
                    }
                    else
                    {
                        MessageBox.Show(responseText);
                    }

                    return;
                }
            }
            MessageBox.Show("No response from marking service");
        }

        class MarkPayload
        {
            public string Function { get; } = "SophisUpdate";
            public MarkParameters Parameters { get; set; } = new MarkParameters();
        }

        class MarkParameters
        {
            public IEnumerable<Mark> list { get; set; } = new List<Mark>() { new Mark()};
        }

        public class Mark {
            public int sicovam { get; set; }
            public string fieldName { get; set; } // Last or T
            public string dateTime { get 
                {
                    var today = DateTime.Today;                    
                    return today.ToString("yyyy-MM-dd");
                }
            }
            public double level { get; set; }
        }


    }
}
