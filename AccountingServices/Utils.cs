using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccountingServices
{
    public static class Utils
    {
        public static bool CaseInsensitiveContains(this string text, string value, StringComparison stringComparison = StringComparison.CurrentCultureIgnoreCase)
        {
            return text.IndexOf(value, stringComparison) >= 0;
        }

        public static FileDate FindLastCacheFile(string cacheDir, string cacheFileNamePrefix)
        {
            string dateFromToRegexPattern = @"(\d{4}\-\d{2}\-\d{2})\-(\d{4}\-\d{2}\-\d{2})\.csv$";
            return FindLastCacheFile(cacheDir, cacheFileNamePrefix, dateFromToRegexPattern, "yyyy-MM-dd", "\\-");
        }

        public static FileDate FindLastCacheFile(string cacheDir, string cacheFileNamePrefix, string dateFromToRegexPattern, string dateParsePattern, string separator)
        {
            var dateDictonary = new SortedDictionary<DateTime, FileDate>();

            string regexp = string.Format("{0}{1}{2}", cacheFileNamePrefix, separator, dateFromToRegexPattern);
            Regex reg = new Regex(regexp);

            string directorySearchPattern = string.Format("{0}*", cacheFileNamePrefix);
            IEnumerable<string> filePaths = Directory.EnumerateFiles(cacheDir, directorySearchPattern);
            foreach (var filePath in filePaths)
            {
                var fileName = Path.GetFileName(filePath);
                var match = reg.Match(fileName);
                if (match.Success)
                {
                    var from = match.Groups[1].Value;
                    var to = match.Groups[2].Value;

                    var dateFrom = DateTime.ParseExact(from, dateParsePattern, CultureInfo.InvariantCulture);
                    var dateTo = DateTime.ParseExact(to, dateParsePattern, CultureInfo.InvariantCulture);
                    var fileDate = new FileDate
                    {
                        From = dateFrom,
                        To = dateTo,
                        FilePath = filePath
                    };
                    dateDictonary.Add(dateTo, fileDate);
                }
            }

            if (dateDictonary.Count() > 0)
            {
                // the first element is the newest date
                return dateDictonary.Last().Value;
            }

            // return a default file date
            return default(FileDate);
        }

        public static List<T> ReadCacheFile<T>(string filePath)
        {
            if (File.Exists(filePath))
            {
                using (TextReader fileReader = File.OpenText(filePath))
                {
                    using (var csvReader = new CsvReader(fileReader))
                    {
                        csvReader.Configuration.Delimiter = ",";
                        csvReader.Configuration.HasHeaderRecord = true;
                        csvReader.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        return csvReader.GetRecords<T>().ToList();
                    }
                }
            }
            else
            {
                return null;
            }
        }

        public static void WriteCacheFile<T>(string filePath, List<T> values)
        {
            using (var sw = new StreamWriter(filePath))
            {
                var csvWriter = new CsvWriter(sw);
                csvWriter.Configuration.Delimiter = ",";
                csvWriter.Configuration.HasHeaderRecord = true;
                csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                csvWriter.WriteRecords(values);
            }
        }

        /// <summary>
        /// Gets the 12:00:00 instance of a DateTime
        /// </summary>
        public static DateTime AbsoluteStart(this DateTime dateTime)
        {
            return dateTime.Date;
        }

        /// <summary>
        /// Gets the 11:59:59 instance of a DateTime
        /// </summary>
        public static DateTime AbsoluteEnd(this DateTime dateTime)
        {
            return AbsoluteStart(dateTime).AddDays(1).AddTicks(-1);
        }

        public static IEnumerable<Tuple<DateTime, DateTime>> SplitDateRange(DateTime start, DateTime end, int dayChunkSize)
        {
            DateTime chunkEnd;
            while ((chunkEnd = start.AddDays(dayChunkSize)) < end)
            {
                yield return Tuple.Create(start, chunkEnd);
                start = chunkEnd;
            }
            yield return Tuple.Create(start, end);
        }
    }

    public class FileDate
    {
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public string FilePath { get; set; }
    }

    public class Date
    {
        DateTime currentDate;
        DateTime yesterday;
        DateTime firstDayOfTheYear;
        DateTime lastDayOfTheYear;

        public DateTime CurrentDate
        {
            get { return Utils.AbsoluteEnd(currentDate); }
        }

        public DateTime Yesterday
        {
            get { return yesterday; }
        }

        public DateTime FirstDayOfTheYear
        {
            get { return firstDayOfTheYear; }
        }

        public DateTime LastDayOfTheYear
        {
            get { return lastDayOfTheYear; }
        }

        public Date()
        {
            currentDate = DateTime.Now.Date;
            yesterday = currentDate.AddDays(-1);
            firstDayOfTheYear = new DateTime(currentDate.Year, 1, 1);
            lastDayOfTheYear = new DateTime(currentDate.Year, 12, 31);
        }
    }
}
