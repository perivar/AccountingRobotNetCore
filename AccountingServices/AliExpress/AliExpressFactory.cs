using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Linq;
using AccountingServices.Helpers;

namespace AccountingServices.AliExpress
{
    public class AliExpressFactory : CachedList<AliExpressOrder>
    {
        public static readonly AliExpressFactory Instance = new AliExpressFactory();

        private AliExpressFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "AliExpress Orders"; } }

        protected override DateTime ForcedUpdateFromDate {
            get
            {
                return new Date().FirstDayOfTheYear;
            }
        }

        public override List<AliExpressOrder> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {        
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding AliExpress Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newAliExpressOrders = AliExpress.ScrapeAliExpressOrders(configuration, from);
            var originalAliExpressOrders = Utils.ReadCacheFile<AliExpressOrder>(lastCacheFileInfo.FilePath);

            // copy all the original AliExpress orders into a new file, except entries that are 
            // from the from date or newer
            var updatedAliExpressOrders = originalAliExpressOrders.Where(p => p.OrderTime < from).ToList();

            // and add the new orders to beginning of list
            updatedAliExpressOrders.InsertRange(0, newAliExpressOrders);

            return updatedAliExpressOrders;
        }

        public override List<AliExpressOrder> GetList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding AliExpress Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return AliExpress.ScrapeAliExpressOrders(configuration, from);
        }
    }
}
