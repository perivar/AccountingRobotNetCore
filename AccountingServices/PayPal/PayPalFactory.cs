using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Linq;
using AccountingServices.Helpers;

namespace AccountingServices.PayPalService
{
    public class PayPalFactory : CachedList<PayPalTransaction>
    {
        public static readonly PayPalFactory Instance = new PayPalFactory();

        private PayPalFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "PayPal Transactions"; } }

        protected override DateTime ForcedUpdateFromDate
        {
            get
            {
                return new Date().FirstDayOfTheYear;
            }
        }

        public override List<PayPalTransaction> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding PayPal transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newPayPalTransactions = Paypal.GetPayPalTransactions(configuration, from, to);
            var originalPayPalTransactions = Utils.ReadCacheFile<PayPalTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original PayPal transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedPayPalTransactions = originalPayPalTransactions.Where(p => p.Timestamp < from).ToList();

            // and add the new transactions to beginning of list
            updatedPayPalTransactions.InsertRange(0, newPayPalTransactions);

            return updatedPayPalTransactions;
        }

        public override List<PayPalTransaction> GetList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding PayPal transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return Paypal.GetPayPalTransactions(configuration, from, to);
        }
    }
}
