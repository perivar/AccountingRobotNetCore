using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Linq;

namespace AccountingServices
{
    public class OberloFactory : CachedList<OberloOrder>
    {
        // get oberlo configuration parameters
        public static readonly OberloFactory Instance = new OberloFactory();

        private OberloFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "Oberlo Orders"; } }

        protected override DateTime ForcedUpdateFromDate
        {
            get
            {
                return new Date().FirstDayOfTheYear;
            }
        }

        public override List<OberloOrder> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding Oberlo Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newOberloOrders = Oberlo.ScrapeOberloOrders(configuration, from, to);
            var originalOberloOrders = Utils.ReadCacheFile<OberloOrder>(lastCacheFileInfo.FilePath);

            // copy all the original Oberlo orders into a new file, except entries that are 
            // from the from date or newer
            var updatedOberloOrders = originalOberloOrders.Where(p => p.CreatedDate < from).ToList();

            // and add the new orders to beginning of list
            updatedOberloOrders.InsertRange(0, newOberloOrders);

            return updatedOberloOrders;
        }

        public override List<OberloOrder> GetList(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding Oberlo Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return Oberlo.ScrapeOberloOrders(configuration, from, to);
        }
    }
}
