using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PayPal.Api;

namespace AccountingServices
{
    public static class Paypal
    {
        public static List<PayPalTransaction> GetPayPalTransactions(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            var payPalTransactions = GetPayPalTransactionsListSoap(configuration, from, to);
            //var payPalTransactions = GetPayPalPaymentListRest(configuration, from, to);

            //var payPalTransactions = new List<PayPalTransaction>();
            //GetPayPalTransactionsListRest(configuration, from, to, payPalTransactions);
            //GetPayPalTransactionsList(configuration, DateTime.Now.AddDays(-31), DateTime.Now, payPalTransactions);

            return payPalTransactions;
        }

        private static List<PayPalTransaction> GetPayPalTransactionsListSoap(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            var payPalTransactions = new List<PayPalTransaction>();

            try
            {
                using (var httpClient = new HttpClient(new HttpClientHandler() { AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip }) { Timeout = TimeSpan.FromSeconds(30) })
                {
                    string payPalApiUsername = configuration.GetValue("PayPalApiUsername");
                    string payPalApiPassword = configuration.GetValue("PayPalApiPassword");
                    string payPalApiSignature = configuration.GetValue("PayPalApiSignature");

                    var soapEnvelopeXml = ConstructSoapEnvelope();
                    var doc = XDocument.Parse(soapEnvelopeXml);

                    var authHeader = doc.Descendants("{urn:ebay:apis:eBLBaseComponents}Credentials").FirstOrDefault();
                    if (authHeader != null)
                    {
                        authHeader.Element("{urn:ebay:apis:eBLBaseComponents}Username").Value = payPalApiUsername;
                        authHeader.Element("{urn:ebay:apis:eBLBaseComponents}Password").Value = payPalApiPassword;
                        authHeader.Element("{urn:ebay:apis:eBLBaseComponents}Signature").Value = payPalApiSignature;
                    }

                    var parameterHeader = doc.Descendants("{urn:ebay:api:PayPalAPI}TransactionSearchRequest").FirstOrDefault();
                    if (parameterHeader != null)
                    {
                        string startDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", from);
                        string endDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", to.AddDays(1));

                        parameterHeader.Element("{urn:ebay:api:PayPalAPI}StartDate").Value = startDate;
                        parameterHeader.Element("{urn:ebay:api:PayPalAPI}EndDate").Value = endDate;
                    }

                    string envelope = doc.ToString();

                    var request = CreateRequest(HttpMethod.Post, "https://api-3t.paypal.com/2.0/", "TransactionSearch", doc);
                    request.Content = new StringContent(envelope, Encoding.UTF8, "text/xml");

                    // request is now ready to be sent via HttpClient
                    HttpResponseMessage response = httpClient.SendAsync(request).Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception();
                    }

                    Task<Stream> streamTask = response.Content.ReadAsStreamAsync();
                    Stream stream = streamTask.Result;
                    var sr = new StreamReader(stream);
                    var soapResponse = XDocument.Load(sr);

                    // parse SOAP response
                    var xmlTransactions = soapResponse.Descendants("{urn:ebay:apis:eBLBaseComponents}PaymentTransactions").ToList();
                    foreach (var xmlTransaction in xmlTransactions)
                    {
                        // build new list
                        var payPalTransaction = new PayPalTransaction();

                        //xmlTransaction.RemoveAttributes(); // trick to ignore ebl types that cannot be deserialized
                        //var transaction = Utils.Deserialize<PaymentTransaction>(xmlTransaction.ToString());

                        payPalTransaction.TransactionID = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}TransactionID").Value;

                        // Converting from paypal date to date:
                        // 2017-08-30T21:13:37Z
                        // var date = DateTimeOffset.Parse(paypalTransaction.Timestamp).UtcDateTime;
                        payPalTransaction.Timestamp = DateTimeOffset.Parse(xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}Timestamp").Value).UtcDateTime;

                        payPalTransaction.Status = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}Status").Value;
                        payPalTransaction.Type = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}Type").Value;

                        payPalTransaction.GrossAmount = decimal.Parse(xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}GrossAmount").Value, CultureInfo.InvariantCulture);
                        payPalTransaction.GrossAmountCurrencyId = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}GrossAmount").Attribute("currencyID").Value;
                        payPalTransaction.NetAmount = decimal.Parse(xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}NetAmount").Value, CultureInfo.InvariantCulture);
                        payPalTransaction.NetAmountCurrencyId = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}NetAmount").Attribute("currencyID").Value;
                        payPalTransaction.FeeAmount = decimal.Parse(xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}FeeAmount").Value, CultureInfo.InvariantCulture);
                        payPalTransaction.FeeAmountCurrencyId = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}FeeAmount").Attribute("currencyID").Value;

                        if (null != xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}Payer"))
                        {
                            payPalTransaction.Payer = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}Payer").Value;
                        }
                        payPalTransaction.PayerDisplayName = xmlTransaction.Element("{urn:ebay:apis:eBLBaseComponents}PayerDisplayName").Value;

                        payPalTransactions.Add(payPalTransaction);
                    }
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine("ERROR: Could not get paypal transactions! '{0}'", e.Message);
            }

            return payPalTransactions;
        }

        private static string ConstructSoapEnvelope()
        {
            var message = @"<?xml version='1.0' encoding='UTF-8'?>
                            <soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:ns='urn:ebay:api:PayPalAPI' xmlns:ebl='urn:ebay:apis:eBLBaseComponents' xmlns:cc='urn:ebay:apis:CoreComponentTypes' xmlns:ed='urn:ebay:apis:EnhancedDataTypes'>
                                <soapenv:Header>
                                    <ns:RequesterCredentials>
                                            <ebl:Credentials>
                                                <ebl:Username></ebl:Username>
                                                <ebl:Password></ebl:Password>
                                                <ebl:Signature></ebl:Signature>
                                            </ebl:Credentials>
                                    </ns:RequesterCredentials>
                                </soapenv:Header>
                                <soapenv:Body>
                                    <ns:TransactionSearchReq>
                                        <ns:TransactionSearchRequest>
                                                <ebl:Version>204.0</ebl:Version>
                                                <ns:StartDate></ns:StartDate>
                                                <ns:EndDate></ns:EndDate>
                                        </ns:TransactionSearchRequest>
                                    </ns:TransactionSearchReq>
                                </soapenv:Body>
                            </soapenv:Envelope>
                            ";
            return message;
        }

        private static HttpRequestMessage CreateRequest(HttpMethod method, string uri, string action, XDocument soapEnvelopeXml)
        {
            var request = new HttpRequestMessage(method: method, requestUri: uri);
            request.Headers.Add("SOAPAction", action);
            request.Headers.Add("ContentType", "text/xml;charset=\"utf-8\"");
            request.Headers.Add("Accept", "text/xml");
            request.Content = new StringContent(soapEnvelopeXml.ToString(), Encoding.UTF8, "text/xml"); ;
            return request;
        }

        private static string GetAccessToken(IMyConfiguration configuration)
        {
            string payPalClientId = configuration.GetValue("PayPalClientId");
            string payPalClientSecret = configuration.GetValue("PayPalClientSecret");

            var config = new Dictionary<string, string>();
            config.Add("mode", PayPal.Api.BaseConstants.LiveMode);
            config.Add("clientId", payPalClientId);
            config.Add("clientSecret", payPalClientSecret);
            config[PayPal.Api.BaseConstants.HttpConnectionTimeoutConfig] = "30000";
            config[PayPal.Api.BaseConstants.HttpConnectionRetryConfig] = "3";

            string accessToken = new PayPal.Api.OAuthTokenCredential(config).GetAccessToken();

            //var apiContext = new PayPal.Api.APIContext(accessToken) { Config = config };
            //var payments = PayPal.Api.Payment.List(apiContext, null, "", null, "", "", DateTime.Now.AddDays(-25).ToString(), DateTime.Now.ToString());

            return accessToken;
        }

        private static HttpClient GetPaypalHttpClient()
        {
            //const string sandbox = "https://api.sandbox.paypal.com";
            const string live = "https://api.paypal.com";

            var http = new HttpClient
            {
                BaseAddress = new Uri(live),
                Timeout = TimeSpan.FromSeconds(30),
            };

            return http;
        }
        private static async Task<PayPalAccessToken> GetPayPalAccessTokenAsync(IMyConfiguration configuration, HttpClient http)
        {
            string payPalClientId = configuration.GetValue("PayPalClientId");
            string payPalClientSecret = configuration.GetValue("PayPalClientSecret");

            byte[] bytes = Encoding.GetEncoding("iso-8859-1").GetBytes($"{payPalClientId}:{payPalClientSecret}");

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "/v1/oauth2/token");
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(bytes));

            var form = new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials"
            };

            request.Content = new FormUrlEncodedContent(form);

            HttpResponseMessage response = await http.SendAsync(request);

            string content = await response.Content.ReadAsStringAsync();
            var accessToken = JsonConvert.DeserializeObject<PayPalAccessToken>(content);
            return accessToken;
        }

        private static void GetPayPalTransactionsListRest(IMyConfiguration configuration, DateTime from, DateTime to, List<PayPalTransaction> payPalTransactions)
        {
            try
            {
                Task.Run(async () =>
                {
                    var httpClient = GetPaypalHttpClient();

                    // Step 1: Get an access token
                    var accessToken = await GetPayPalAccessTokenAsync(configuration, httpClient);

                    // split the date range into smaller chunks since the maximum number of days in the range supported is 31
                    var dateRanges = Utils.SplitDateRange(from, to, 31);

                    foreach (var dateRange in dateRanges)
                    {
                        // Step 2: Get the transactions
                        await GetPayPalTransactionsAsync(httpClient, accessToken, dateRange.Item1, dateRange.Item2, payPalTransactions);
                    }

                }).Wait();
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Could not get paypal transactions! '{0}'", e.Message);
            }
        }
        private static async Task GetPayPalTransactionsAsync(HttpClient httpClient, PayPalAccessToken accessToken, DateTime from, DateTime to, List<PayPalTransaction> payPalTransactions)
        {
            string startDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", from);
            string endDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", to);
            string url = $"https://api.paypal.com/v1/reporting/transactions?fields=all&page_size=100&page=1&start_date={startDate}&end_date={endDate}";

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.access_token);
            HttpResponseMessage response = await httpClient.SendAsync(request);
            string jsonString = await response.Content.ReadAsStringAsync();
            dynamic jsonD = JsonConvert.DeserializeObject(jsonString);

            // parse json 
            string accountNumber = jsonD.account_number;
            int total_pages = jsonD.total_pages;
            DateTime dateStart = jsonD.start_date;
            DateTime dateEnd = jsonD.end_date;
            int page = jsonD.page;
            int totalItems = jsonD.total_items;

            foreach (var transaction in jsonD.transaction_details)
            {
                var transactionInfo = transaction.transaction_info;
                var transactionId = transactionInfo.transaction_id;
                string transactionEventCode = transactionInfo.transaction_event_code;
                DateTime transactionInitiationDate = transactionInfo.transaction_initiation_date;
                DateTime transactionUpdatedDate = transactionInfo.transaction_updated_date;
                var transactionAmountObject = transactionInfo.transaction_amount;
                string transactionAmountCurrencyCode = transactionAmountObject.currency_code;
                decimal transactionAmountValue = transactionAmountObject.value;
                string transactionStatus = transactionInfo.transaction_status;
                string bankReferenceId = transactionInfo.bank_reference_id;
                var endingBalanceObject = transactionInfo.ending_balance;
                string endingBalanceCurrencyCode = endingBalanceObject.currency_code;
                decimal endingBalanceValue = endingBalanceObject.value;
                var availableBalanceObject = transactionInfo.available_balance;
                string availableBalanceCurrencyCode = availableBalanceObject.currency_code;
                decimal availableBalanceValue = availableBalanceObject.value;
                string protectionEligibility = transactionInfo.protection_eligibility;

                PayPalTransaction payPalTransaction = new PayPalTransaction();
                payPalTransaction.TransactionID = transactionId;
                payPalTransaction.Timestamp = transactionInitiationDate;
                payPalTransaction.Status = transactionStatus;
                payPalTransaction.Type = transactionEventCode;
                payPalTransaction.GrossAmount = transactionAmountValue;
                payPalTransaction.NetAmount = 0;
                payPalTransaction.FeeAmount = 0;

                // payPalTransaction.Payer = transaction.Payer;
                // payPalTransaction.PayerDisplayName = transaction.PayerDisplayName;

                payPalTransactions.Add(payPalTransaction);
            }
        }

        private static List<PayPalTransaction> GetPayPalPaymentListRest(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            // NOTE THIS WILL ALWAYS RETURN ZERO ELEMENTS
            // since only transactions created with the api will be returned
            var payPalTransactions = new List<PayPalTransaction>();
            try
            {
                Task.Run(async () =>
                {
                    var httpClient = GetPaypalHttpClient();

                    // Step 1: Get an access token
                    var accessToken = await GetPayPalAccessTokenAsync(configuration, httpClient);

                    // split the date range into smaller chunks since the maximum number of days in the range supported is 31
                    var dateRanges = Utils.SplitDateRange(from, to, 31);

                    foreach (var dateRange in dateRanges)
                    {
                        // Step 2: Get the transactions
                        await GetPayPalPaymentListAsync(httpClient, accessToken, dateRange.Item1, dateRange.Item2, payPalTransactions);
                    }

                }).Wait();
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Could not get paypal transactions! '{0}'", e.Message);
            }
            return payPalTransactions;
        }
        private static async Task GetPayPalPaymentListAsync(HttpClient httpClient, PayPalAccessToken accessToken, DateTime from, DateTime to, List<PayPalTransaction> payPalTransactions)
        {
            // https://developer.paypal.com/docs/integration/direct/payments/paypal-payments/?mark=list%20payment#search-payment-details
            string startTime = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", from);
            string endTime = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", to);
            string url = $"https://api.paypal.com/v1/payments/payment?count=100&start_time={startTime}&end_time={endTime}";

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.access_token);
            HttpResponseMessage response = await httpClient.SendAsync(request);
            string jsonString = await response.Content.ReadAsStringAsync();
            dynamic jsonD = JsonConvert.DeserializeObject(jsonString);

            // parse json 
            //payPalTransactions.Add(payPalTransaction);
        }
    }

    class PayPalAccessToken
    {
        public string scope { get; set; }
        public string nonce { get; set; }
        public string access_token { get; set; }
        public string token_type { get; set; }
        public string app_id { get; set; }
        public long expires_in { get; set; }
    }

    [XmlRoot(ElementName = "PaymentTransactions", Namespace = "urn:ebay:apis:eBLBaseComponents")]
    public class PaymentTransaction
    {
        [XmlElement("Timestamp")]
        public string Timestamp { get; set; }

        [XmlElement("Timezone")]
        public string Timezone { get; set; }

        [XmlElement("Type")]
        public string Type { get; set; }

        [XmlElement("Payer")]
        public string Payer { get; set; }

        [XmlElement("PayerDisplayName")]
        public string PayerDisplayName { get; set; }

        [XmlElement("TransactionID")]
        public string TransactionID { get; set; }

        [XmlElement("Status")]
        public string Status { get; set; }

        [XmlElement("GrossAmount")]
        public string GrossAmount { get; set; }

        [XmlElement("FeeAmount")]
        public string FeeAmount { get; set; }

        [XmlElement("NetAmount")]
        public string NetAmount { get; set; }
    }
}
