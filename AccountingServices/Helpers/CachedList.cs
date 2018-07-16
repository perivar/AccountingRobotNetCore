using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace AccountingServices.Helpers
{
    public abstract class CachedList<T>
    {
        protected abstract string CacheFileNamePrefix { get; }

        protected abstract DateTime ForcedUpdateFromDate { get; }

        public List<T> GetLatest(IMyConfiguration configuration, TextWriter writer, bool forceUpdate = false)
        {
            var date = new Date();
            var currentDate = date.CurrentDate;
            var firstDayOfTheYear = date.FirstDayOfTheYear;

            // default is a lookup from beginning of year until now
            DateTime from = firstDayOfTheYear;
            DateTime to = currentDate;

            string cacheDir = configuration.GetValue("CacheDir");

            if (forceUpdate)
            {
                writer.WriteLine("Forcing updating from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", ForcedUpdateFromDate, to);
                var values = GetList(configuration, writer, ForcedUpdateFromDate, to);

                string forcedCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, ForcedUpdateFromDate, to));
                Utils.WriteCacheFile(forcedCacheFilePath, values);
                writer.WriteLine("Successfully wrote file to {0}", forcedCacheFilePath);
                return values;
            }

            // check if we have a cache file
            var lastCacheFileInfo = Utils.FindLastCacheFile(cacheDir, CacheFileNamePrefix);

            // if the cache file object has values
            if (lastCacheFileInfo != null && !lastCacheFileInfo.Equals(default(FileDate)))
            {
                // find values starting from when the cache file ends and until now
                from = lastCacheFileInfo.To;
                to = currentDate;

                // if the from date is today, then we already have an updated file so use cache
                if (lastCacheFileInfo.To.Date.Equals(currentDate.Date))
                {
                    // use latest cache file (or update if the cache file is empty)
                    return GetList(configuration, writer, lastCacheFileInfo.FilePath, from, to);
                }
                else if (from != firstDayOfTheYear)
                {
                    // combine new and old values
                    var updatedValues = GetCombinedUpdatedAndExisting(configuration, writer, lastCacheFileInfo, from, to);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, firstDayOfTheYear, to));
                    Utils.WriteCacheFile(newCacheFilePath, updatedValues);
                    writer.WriteLine("Successfully wrote file to {0}", newCacheFilePath);
                    return updatedValues;
                }
            }

            // get updated transactions (or from cache file)
            string cacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, from, to));
            return GetList(configuration, writer, cacheFilePath, from, to);
        }

        public List<T> GetList(IMyConfiguration configuration, TextWriter writer, string cacheFilePath, DateTime from, DateTime to)
        {
            var cachedList = Utils.ReadCacheFile<T>(cacheFilePath);
            //if (cachedList != null && cachedList.Count() > 0)
            if (cachedList != null)
            {
                writer.WriteLine("Using cache file {0}.", cacheFilePath);
                return cachedList;
            }
            else
            {
                writer.WriteLine("Cache file is empty. Updating ...");
                var values = GetList(configuration, writer, from, to);
                Utils.WriteCacheFile(cacheFilePath, values);
                writer.WriteLine("Successfully wrote file to {0}", cacheFilePath);
                return values;
            }
        }

        public abstract List<T> GetList(IMyConfiguration configuration, TextWriter writer, DateTime from, DateTime to);

        public abstract List<T> GetCombinedUpdatedAndExisting(IMyConfiguration configuration, TextWriter writer, FileDate lastCacheFileInfo, DateTime from, DateTime to);
    }
}
