using CsvHelper;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

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
            var dateDictionary = new SortedDictionary<DateTime, FileDate>();

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
                    dateDictionary.Add(dateTo, fileDate);
                }
            }

            if (dateDictionary.Count() > 0)
            {
                // the first element is the newest date
                return dateDictionary.Last().Value;
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

        public static T Deserialize<T>(string xmlStr)
        {
            var serializer = new XmlSerializer(typeof(T));
            T result;
            using (TextReader reader = new StringReader(xmlStr))
            {
                result = (T)serializer.Deserialize(reader);
            }
            return result;
        }

        /// <summary>
        /// Find the path relative to the running assembly
        /// Use like this Utils.GetFilePathRelativeToAssembly(@"..\..\..\..\AccountingServices\bin\debug\netcoreapp2.0");
        /// </summary>
        /// <param name="pathRelativeToAssembly">relative path</param>
        /// <returns>the full path relative to the assembly</returns>
        public static string GetFilePathRelativeToAssembly(string pathRelativeToAssembly)
        {
            string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string filePathRelativeToAssembly = Path.Combine(assemblyPath, pathRelativeToAssembly);
            string normalizedPath = Path.GetFullPath(filePathRelativeToAssembly);
            return normalizedPath;
        }

        public static IWebDriver GetChromeWebDriver(string userDataDir, string chromeDriverExePath)
        {

            // workaround to a bug in dot net core that makes findelement so slow
            // https://github.com/SeleniumHQ/selenium/issues/4988
            // change the chrome driver to run on another port and 127.0.0.1 instead of localhost

            // add chromedriver to the PATH
            var chromeDriverDirectory = new FileInfo(chromeDriverExePath).FullName;
            string pathEnv = Environment.GetEnvironmentVariable("PATH");
            pathEnv += ";" + chromeDriverDirectory;
            Environment.SetEnvironmentVariable("PATH", pathEnv);

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.Port = 5555; // Any port value.
            service.Start();

            ChromeOptions options = new ChromeOptions();
            string userDataArgument = string.Format("user-data-dir={0}", userDataDir);
            options.AddArguments(userDataArgument);
            options.AddArguments("--start-maximized");
            options.AddArgument("--log-level=3");
            //options.AddArguments("--ignore-certificate-errors");
            //options.AddArguments("--ignore-ssl-errors");
            //options.AddArgument("--headless");

            //IWebDriver driver = new ChromeDriver(chromeDriverExePath, options);
            IWebDriver driver = new RemoteWebDriver(new Uri("http://127.0.0.1:5555"), options);
            return driver;
        }

        public static string CreateMD5(string input)
        {
            // byte array representation of that string
            byte[] inputBytes = new UTF8Encoding().GetBytes(input);

            // need MD5 to calculate the hash
            byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(inputBytes);

            // string representation (similar to UNIX format)
            string encoded = BitConverter.ToString(hash)
               // without dashes
               .Replace("-", string.Empty)
               // make lowercase
               .ToLower();

            return encoded;
        }

        public static decimal CurrencyConverter(string fromCurrency, string toCurrency)
        {
            string query = String.Format("{0}_{1}", WebUtility.UrlEncode(fromCurrency), WebUtility.UrlEncode(toCurrency));
            string url = String.Format("https://free.currencyconverterapi.com/api/v5/convert?q={0}&compact=ultra", query);

            using (var client = new WebClient())
            {
                // make sure we read in utf8
                client.Encoding = System.Text.Encoding.UTF8;

                string json = client.DownloadString(url);

                // parse json
                dynamic jsonObject = JsonConvert.DeserializeObject(json);
                if (jsonObject != null)
                {
                    //var value = (decimal)jsonObject["results"][query]["val"]; // if using compact view
                    var value = (decimal)jsonObject[query]; // if using ultra compact view
                    return value;
                }
            }
            return 0;
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
