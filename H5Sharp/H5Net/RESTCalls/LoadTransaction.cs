using H5Net.JsonResponse;
using H5Net.ReadAPIObjects;
using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace H5Net.RESTCalls
{
    public class LoadTransaction
    {
        public LoadTransaction()
        {
        }

        //TODO: Check if memory streams can be used here. Verify if less memory is used via memory streams
        public static async Task<Response> LoadBulkJSONTransaction(string jsonPayLoad, string baseAddress, string apiEndPoint, string accessToken)
        {
            //var apiTransOutput = "";

            string payLoad = jsonPayLoad; //IONReader.ReadJSONPayLoad(jsonPayLoad);
            var client = new HttpClient(
                new HttpClientHandler()
                {
                    AutomaticDecompression = DecompressionMethods.GZip
                });
            client.BaseAddress = new Uri(baseAddress);
            //{ BaseAddress = new Uri(baseAddress) };
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.AcceptEncoding.Add(new StringWithQualityHeaderValue("gzip"));

            var content = new StringContent(payLoad, Encoding.UTF8, "application/json");
            client.SetBearerToken(accessToken);

            //  var response = client.PostAsync(apiEndPoint, content).Result;
            HttpResponseMessage response = await client.PostAsync(apiEndPoint, content).ConfigureAwait(false);
            //var res = await response.Content.ReadAsStringAsync

            /*
             * if (response.IsSuccessStatusCode)
            {
                apiTransOutput = response.Content.ReadAsStringAsync().Result;
            }
            */

           var apiTransOutput = response.Content.ReadAsStringAsync();

            var message = JsonConvert.DeserializeObject<Response>(apiTransOutput.Result);
            
            // var message = JsonConvert.DeserializeObject(apiTransOutput);
            return message;
        }

        //TODO: Implement generic version of Lst transaction. Might need a generic version for Get transaction as well
        public static async Task<AgreementLines> ExecuteLstTransaction(string jsonPayLoad, string baseAddress, string apiEndPoint, string accessToken, string pgmName, string transName)
        {
            //var apiTransOutput = "";

            string payLoad = jsonPayLoad; //IONReader.ReadJSONPayLoad(jsonPayLoad);
            var client = new HttpClient(
                new HttpClientHandler()
                {
                    AutomaticDecompression = DecompressionMethods.GZip
                });
            client.BaseAddress = new Uri(baseAddress);

            //{ BaseAddress = new Uri(baseAddress) };
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.AcceptEncoding.Add(new StringWithQualityHeaderValue("gzip"));

            var content = new StringContent(payLoad, Encoding.UTF8, "application/json");
            client.SetBearerToken(accessToken);

            //var response = client.PostAsync(apiEndPoint, content).Result;

            HttpResponseMessage response = await client.PostAsync(apiEndPoint, content).ConfigureAwait(false);

            /*
             * if (response.IsSuccessStatusCode)
            {
                apiTransOutput = response.Content.ReadAsStringAsync().Result;
            }
            */

            var apiTransOutput = response.Content.ReadAsStringAsync();

            var AgreementLines = JsonConvert.DeserializeObject<AgreementLines>(apiTransOutput.Result);
            return AgreementLines;
        }

        public static async Task<PurchaseOrderLines> ExecutePurchaseLstTransaction(string jsonPayLoad, string baseAddress, string apiEndPoint, string accessToken, string pgmName, string transName)
        {
            //var apiTransOutput = "";

            string payLoad = jsonPayLoad; //IONReader.ReadJSONPayLoad(jsonPayLoad);
            var client = new HttpClient(
                new HttpClientHandler()
                {
                    AutomaticDecompression = DecompressionMethods.GZip
                });
            client.BaseAddress = new Uri(baseAddress);

            //{ BaseAddress = new Uri(baseAddress) };
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.AcceptEncoding.Add(new StringWithQualityHeaderValue("gzip"));

            var content = new StringContent(payLoad, Encoding.UTF8, "application/json");
            client.SetBearerToken(accessToken);

            // var response = client.PostAsync(apiEndPoint, content).Result;
            HttpResponseMessage response = await client.PostAsync(apiEndPoint, content).ConfigureAwait(false);

            /*
             * if (response.IsSuccessStatusCode)
            {
                apiTransOutput = response.Content.ReadAsStringAsync().Result;
            }
            */

            var apiTransOutput = response.Content.ReadAsStringAsync();

            var PurchaseOrderLines = JsonConvert.DeserializeObject<PurchaseOrderLines>(apiTransOutput.Result);
            return PurchaseOrderLines;
        }
    }
}