using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Linq;
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

        public override List<StripeTransaction> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newStripeTransactions = Stripe.GetStripeChargeTransactions(configuration, from, to);
            var originalStripeTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original stripe transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedStripeTransactions = originalStripeTransactions.Where(p => p.Created < from).ToList();

            // and add the new transactions to beginning of list
            updatedStripeTransactions.InsertRange(0, newStripeTransactions);

            return updatedStripeTransactions;
        }

        public override List<StripeTransaction> GetList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return Stripe.GetStripeChargeTransactions(configuration, from, to);
        }
    }
}
