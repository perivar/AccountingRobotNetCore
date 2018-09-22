using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Http;
using System.Globalization;
using IdentityModel.Client;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AccountingServices.Helpers;

namespace AccountingServices.SBankenService
{
    public class SBankenFactory : CachedList<SBankenTransaction>
    {
        public static readonly SBankenFactory Instance = new SBankenFactory();

        private SBankenFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "SBanken Transactions"; } }

        protected override DateTime ForcedUpdateFromDate
        {
            get
            {
                return new Date().FirstDayOfTheYear;
            }
        }

        public override async Task<List<SBankenTransaction>> GetCombinedUpdatedAndExistingAsync(IMyConfiguration configuration, TextWriter writer, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            await writer.WriteLineAsync(string.Format("Finding SBanken transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to));
            var newSBankenTransactions = await GetSBankenTransactionsAsync(configuration, from, to);
            var originalSBankenTransactions = Utils.ReadCacheFile<SBankenTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original SBanken transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedSBankenTransactions = originalSBankenTransactions.Where(p => p.TransactionDate < from).ToList();

            // and add the new transactions to beginning of list
            updatedSBankenTransactions.InsertRange(0, newSBankenTransactions);

            return updatedSBankenTransactions;
        }

        public override async Task<List<SBankenTransaction>> GetListAsync(IMyConfiguration configuration, TextWriter writer, DateTime from, DateTime to)
        {
            await writer.WriteLineAsync(string.Format("Finding SBanken transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to));
            return await GetSBankenTransactionsAsync(configuration, from, to);
        }

        private async Task<List<SBankenTransaction>> GetSBankenTransactionsAsync(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            // get SBanken configuration parameters
            string clientId = configuration.GetValue("SBankenApiClientId");
            string secret = configuration.GetValue("SBankenApiSecret");
            string customerId = configuration.GetValue("SBankenApiCustomerId");
            string accountId = configuration.GetValue("SBankenAccountId");

            var sBankenTransactions = new List<SBankenTransaction>();

            // see also
            // https://github.com/anderaus/Sbanken.DotNet/blob/master/src/Sbanken.DotNet/Http/Connection.cs

            /** Setup constants */
            const string discoveryEndpoint = "https://auth.sbanken.no/identityserver";
            const string apiBaseAddress = "https://api.sbanken.no";
            const string bankBasePath = "/bank";

            // First: get the OpenId configuration from Sbanken.
            var discoClient = new DiscoveryClient(discoveryEndpoint);

            var x = discoClient.Policy = new DiscoveryPolicy()
            {
                ValidateIssuerName = false,
            };

            var discoResult = await discoClient.GetAsync();

            if (discoResult.Error != null)
            {
                throw new Exception(discoResult.Error);
            }

            // The application now knows how to talk to the token endpoint.

            // Second: the application authenticates against the token endpoint
            // ensure basic authentication RFC2617 is used
            // The application must authenticate itself with Sbanken's authorization server.
            // The basic authentication scheme is used here (https://tools.ietf.org/html/rfc2617#section-2 ) 
            //var tokenClient = new TokenClient(discoResult.TokenEndpoint, clientId, secret)
            //{
            //    BasicAuthenticationHeaderStyle = BasicAuthenticationHeaderStyle.Rfc2617
            //};
            var tokenClient = new TokenClient(discoResult.TokenEndpoint, clientId, secret);

            var tokenResponse = await tokenClient.RequestClientCredentialsAsync();

            if (tokenResponse.IsError)
            {
                throw new Exception(tokenResponse.ErrorDescription);
            }

            // The application now has an access token.

            var httpClient = new HttpClient()
            {
                BaseAddress = new Uri(apiBaseAddress),
                DefaultRequestHeaders =
                {
                    { "customerId", customerId }
                }
            };

            // Finally: Set the access token on the connecting client. 
            // It will be used with all requests against the API endpoints.
            httpClient.SetBearerToken(tokenResponse.AccessToken);

            // retrieve the customer's transactions
            // RFC3339 / ISO8601 with 3 decimal places
            // yyyy-MM-ddTHH:mm:ss.fffK            
            string querySuffix = string.Format(CultureInfo.InvariantCulture, "?length=1000&startDate={0:yyyy-MM-ddTHH:mm:ss.fffK}&endDate={1:yyyy-MM-ddTHH:mm:ss.fffK}", from, to);
            var transactionResponse = await httpClient.GetAsync($"{bankBasePath}/api/v1/Transactions/{accountId}{querySuffix}");
            var transactionResult = await transactionResponse.Content.ReadAsStringAsync();

            // parse json
            dynamic jsonDe = JsonConvert.DeserializeObject(transactionResult);

            if (jsonDe != null)
            {
                foreach (var transaction in jsonDe.items)
                {
                    var amount = transaction.amount;
                    var text = transaction.text;
                    var transactionType = transaction.transactionType;
                    var transactionTypeText = transaction.transactionTypeText;
                    var accountingDate = transaction.accountingDate;
                    var interestDate = transaction.interestDate;

                    var transactionId = transaction.transactionId;
                    // Note, until Sbanken fixed their unique transaction Id issue, generate one ourselves
                    if (transactionId == null || !transactionId.HasValues || transactionId == JTokenType.Null)
                    {
                        // ensure we are working with proper objects
                        DateTime localAccountingDate = accountingDate;
                        DateTime localInterestDate = interestDate;
                        decimal localAmount = amount;

                        // convert to string
                        string accountingDateString = localAccountingDate.ToString("G", CultureInfo.InvariantCulture);
                        string interestDateString = localInterestDate.ToString("G", CultureInfo.InvariantCulture);
                        string amountString = localAmount.ToString("G", CultureInfo.InvariantCulture);

                        // combine and create md5
                        string uniqueContent = $"{accountingDateString}{interestDateString}{transactionTypeText}{text}{amountString}";
                        string hashCode = Utils.CreateMD5(uniqueContent);
                        transactionId = hashCode;
                    }

                    var sBankenTransaction = new SBankenTransaction();
                    sBankenTransaction.TransactionDate = accountingDate;
                    sBankenTransaction.InterestDate = interestDate;
                    sBankenTransaction.ArchiveReference = transactionId;
                    sBankenTransaction.Type = transactionTypeText;
                    sBankenTransaction.Text = text;

                    var date = new Date();
                    var currentDate = date.CurrentDate;

                    // check if card details was specified
                    if ((bool)transaction.cardDetailsSpecified)
                    {
                        var cardDatailsCardNumber = transaction.cardDetails.cardNumber;
                        var cardDatailsCurrencyAmount = transaction.cardDetails.currencyAmount;
                        var cardDatailsCurrencyRate = transaction.cardDetails.currencyRate;
                        var cardDatailsCurrencyCode = transaction.cardDetails.originalCurrencyCode;
                        var cardDatailsMerchantCategoryCode = transaction.cardDetails.merchantCategoryCode;
                        var cardDatailsMerchantName = transaction.cardDetails.merchantName;
                        var cardDatailsPurchaseDate = transaction.cardDetails.purchaseDate;
                        var cardDatailsTransactionId = transaction.cardDetails.transactionId;

                        sBankenTransaction.ExternalPurchaseDate = cardDatailsPurchaseDate;
                        sBankenTransaction.ExternalPurchaseAmount = cardDatailsCurrencyAmount;
                        sBankenTransaction.ExternalPurchaseCurrency = cardDatailsCurrencyCode;
                        sBankenTransaction.ExternalPurchaseVendor = cardDatailsMerchantName;
                        sBankenTransaction.ExternalPurchaseExchangeRate = cardDatailsCurrencyRate;

                        // NOTE! fix a likely bug in the API where the external purchase date is the wrong year
                        if (sBankenTransaction.ExternalPurchaseDate.Year == currentDate.Year
                            && sBankenTransaction.ExternalPurchaseDate.Month > currentDate.Month)
                        {
                            sBankenTransaction.ExternalPurchaseDate = sBankenTransaction.ExternalPurchaseDate.AddYears(-1);
                        }
                    }

                    // set account change
                    sBankenTransaction.AccountChange = amount;

                    if (amount > 0)
                    {
                        sBankenTransaction.AccountingType = SBankenTransaction.AccountingTypeEnum.IncomeUnknown;
                    }
                    else
                    {
                        sBankenTransaction.AccountingType = SBankenTransaction.AccountingTypeEnum.CostUnknown;
                    }

                    if (transactionId != null && transactionId != "")
                    {
                        sBankenTransactions.Add(sBankenTransaction);
                    }
                }
            }

            return sBankenTransactions;
        }
    }
}
