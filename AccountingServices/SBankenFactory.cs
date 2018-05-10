using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using IdentityModel.Client;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Globalization;

namespace AccountingServices
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

        public override List<SBankenTransaction> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding SBanken transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newSBankenTransactions = GetSBankenTransactions(configuration, from, to);
            var originalSBankenTransactions = Utils.ReadCacheFile<SBankenTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original SBanken transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedSBankenTransactions = originalSBankenTransactions.Where(p => p.TransactionDate < from).ToList();

            // and add the new transactions to beginning of list
            updatedSBankenTransactions.InsertRange(0, newSBankenTransactions);

            return updatedSBankenTransactions;
        }

        public override List<SBankenTransaction> GetList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding SBanken transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return GetSBankenTransactions(configuration, from, to);
        }

        private List<SBankenTransaction> GetSBankenTransactions(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            try
            {
                return GetSBankenTransactionsAsync(configuration, from, to).GetAwaiter().GetResult();
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Could not get transactions from SBanken! '{0}'", e.Message);
                return new List<SBankenTransaction>();
            }
        }

        private async Task<List<SBankenTransaction>> GetSBankenTransactionsAsync(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            // get SBanken configuration parameters
            string clientId = configuration.GetValue("SBankenApiClientId");
            string secret = configuration.GetValue("SBankenApiSecret");
            string customerId = configuration.GetValue("SBankenApiCustomerId");
            string accountNumber = configuration.GetValue("SBankenAccountNumber");

            var sBankenTransactions = new List<SBankenTransaction>();

            // see also
            // https://github.com/anderaus/Sbanken.DotNet/blob/master/src/Sbanken.DotNet/Http/Connection.cs

            /** Setup constants */
            const string discoveryEndpoint = "https://api.sbanken.no/identityserver";
            const string apiBaseAddress = "https://api.sbanken.no";
            const string bankBasePath = "/bank";
            //const string customersBasePath = "/customers";

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
            var tokenClient = new TokenClient(discoResult.TokenEndpoint, clientId, secret)
            {
                BasicAuthenticationHeaderStyle = BasicAuthenticationHeaderStyle.Rfc2617
            };

            var tokenResponse = tokenClient.RequestClientCredentialsAsync().Result;

            if (tokenResponse.IsError)
            {
                throw new Exception(tokenResponse.Error);
            }

            // The application now has an access token.

            var httpClient = new HttpClient()
            {
                BaseAddress = new Uri(apiBaseAddress),
            };

            // Finally: Set the access token on the connecting client. 
            // It will be used with all requests against the API endpoints.
            httpClient.SetBearerToken(tokenResponse.AccessToken);

            // retrieves the customer's information.
            //var customerResponse = await httpClient.GetAsync($"{customersBasePath}/api/v1/Customers/{customerId}");
            //var customerResult = await customerResponse.Content.ReadAsStringAsync();

            // retrieves the customer's accounts.
            //var accountResponse = await httpClient.GetAsync($"{bankBasePath}/api/v1/Accounts/{customerId}");
            //var accountResult = await accountResponse.Content.ReadAsStringAsync();

            // retrieve the customer's transactions
            // RFC3339 / ISO8601 with 3 decimal places
            // yyyy-MM-ddTHH:mm:ss.fffK            
            string querySuffix = string.Format(CultureInfo.InvariantCulture, "?startDate={0:yyyy-MM-ddTHH:mm:ss.fffK}&endDate={1:yyyy-MM-ddTHH:mm:ss.fffK}", from, to);
            //var transactionResponse = await httpClient.GetAsync($"{bankBasePath}/api/v1/Transactions/{customerId}/{accountNumber}{querySuffix}");
            var transactionResponse = await httpClient.GetAsync($"{bankBasePath}/api/v2/Transactions/{customerId}/{accountNumber}{querySuffix}");
            var transactionResult = await transactionResponse.Content.ReadAsStringAsync();

            // parse json
            dynamic jsonDe = JsonConvert.DeserializeObject(transactionResult);

            foreach (var transaction in jsonDe.items)
            {
                //var transactionId = transaction.transactionId;
                var amount = transaction.amount;
                var text = transaction.text;
                var transactionType = transaction.transactionType;
                var transactionTypeText = transaction.transactionTypeText;
                var accountingDate = transaction.accountingDate;
                var interestDate = transaction.interestDate;

                // Note, untill Sbanken fixed their unique transaction Id issue, generate one ourselves
                string uniqueContent = $"{accountingDate}{interestDate}{text}{amount}";
                string hashCode = String.Format("{0:X}", uniqueContent.GetHashCode());
                var transactionId = hashCode;

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

            return sBankenTransactions;
        }
    }
}
