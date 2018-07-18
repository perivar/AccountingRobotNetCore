using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using AccountingServices.Helpers;

namespace AccountingServices.StripeService
{
    public class StripeChargeFactory : CachedList<StripeTransaction>
    {
        public static readonly StripeChargeFactory Instance = new StripeChargeFactory();

        private StripeChargeFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "Stripe Transactions"; } }

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
            await writer.WriteLineAsync(string.Format("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to));
            var newStripeTransactions = await Stripe.GetStripeChargeTransactionsAsync(configuration, from, to);
            var originalStripeTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original stripe transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedStripeTransactions = originalStripeTransactions.Where(p => p.Created < from).ToList();

            // and add the new transactions to beginning of list
            updatedStripeTransactions.InsertRange(0, newStripeTransactions);

            return updatedStripeTransactions;
        }

        public override async Task<List<StripeTransaction>> GetListAsync(IMyConfiguration configuration, TextWriter writer, DateTime from, DateTime to)
        {
            await writer.WriteLineAsync(string.Format("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to));
            return await Stripe.GetStripeChargeTransactionsAsync(configuration, from, to);
        }
    }
}
