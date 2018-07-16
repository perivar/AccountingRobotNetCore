using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Linq;
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

        public override List<StripeTransaction> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newStripePayoutTransactions = Stripe.GetStripePayoutTransactions(configuration, from, to);
            var originalStripePayoutTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original stripe transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedStripePayoutTransactions = originalStripePayoutTransactions.Where(p => p.Created < from).ToList();

            // and add the new transactions to beginning of list
            updatedStripePayoutTransactions.InsertRange(0, newStripePayoutTransactions);

            return updatedStripePayoutTransactions;
        }

        public override List<StripeTransaction> GetList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return Stripe.GetStripePayoutTransactions(configuration, from, to);
        }
    }
}
