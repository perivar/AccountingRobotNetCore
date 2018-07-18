using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using AccountingServices.Helpers;

namespace AccountingServices.StripeService
{
    public class StripePayoutFactory : CachedList<StripeTransaction>
    {
        public static readonly StripePayoutFactory Instance = new StripePayoutFactory();

        private StripePayoutFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "Stripe Payout Transactions"; } }

        protected override DateTime ForcedUpdateFromDate
        {
            get
            {
                return new Date().FirstDayOfTheYear;
            }
        }

        public override async Task<List<StripeTransaction>> GetCombinedUpdatedAndExistingAsync(IMyConfiguration configuration, TextWriter writer, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            await writer.WriteLineAsync(string.Format("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to));
            var newStripePayoutTransactions = await Stripe.GetStripePayoutTransactionsAsync(configuration, from, to);
            var originalStripePayoutTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original stripe transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedStripePayoutTransactions = originalStripePayoutTransactions.Where(p => p.Created < from).ToList();

            // and add the new transactions to beginning of list
            updatedStripePayoutTransactions.InsertRange(0, newStripePayoutTransactions);

            return updatedStripePayoutTransactions;
        }

        public override async Task<List<StripeTransaction>> GetListAsync(IMyConfiguration configuration, TextWriter writer, DateTime from, DateTime to)
        {
            await writer.WriteLineAsync(string.Format("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to));
            return await Stripe.GetStripePayoutTransactionsAsync(configuration, from, to);
        }
    }
}
