using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PayPal.Api;

namespace AccountingServices
{
    public static class Paypal
    {
        public static List<PayPalTransaction> GetPayPalTransactions(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            //var transaction = GetPayPalTransactionsList(configuration, from, to);
            return GetPayPalTransactionsList(configuration, DateTime.Now.AddDays(-31), DateTime.Now);
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

        private static List<PayPalTransaction> GetPayPalTransactionsList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            var payPalTransactions = new List<PayPalTransaction>();

            try
            {
                Task.Run(async () =>
                {
                    var httpClient = GetPaypalHttpClient();

                    // Step 1: Get an access token
                    var accessToken = await GetPayPalAccessTokenAsync(configuration, httpClient);

                    // Step 2: List the transactions
                    payPalTransactions = await GetPayPalTransactionsAsync(httpClient, accessToken, from, to);
                }).Wait();
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Could not get paypal transactions! '{0}'", e.Message);
            }

            return payPalTransactions;
        }
        private static async Task<List<PayPalTransaction>> GetPayPalTransactionsAsync(HttpClient httpClient, PayPalAccessToken accessToken, DateTime from, DateTime to)
        {
            var payPalTransactions = new List<PayPalTransaction>();

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

            return payPalTransactions;
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
}
