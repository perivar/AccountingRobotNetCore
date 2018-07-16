using ClosedXML.Excel;
using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AccountingServices.Excel
{
    public static class ExcelUtils
    {
        /// <summary>
        /// Convert from excel date int string to actual date
        /// E.g. 39938 gets converted to 05/05/2009
        /// </summary>
        /// <param name="dateIntString">int string like 39938</param>
        /// <returns>datetime object (like 05/05/2009)</returns>
        public static DateTime GetDateFromExcelDateInt(string dateIntString)
        {
            try
            {
                double d = double.Parse(dateIntString);
                DateTime dateTime = DateTime.FromOADate(d);
                return dateTime;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Convert from excel decimal string to a decimal
        /// </summary>
        /// <param name="currencyString">currency string like 133.3</param>
        /// <returns>decimal like 133.3</returns>
        public static decimal GetDecimalFromExcelCurrencyString(string currencyString, bool isNorwegian = false)
        {
            if (isNorwegian)
            {
                return Convert.ToDecimal(currencyString, CultureInfo.GetCultureInfo("no"));
            }
            else
            {
                return Convert.ToDecimal(currencyString, CultureInfo.InvariantCulture);
            }
        }

        public static DateTime GetDateFromBankStatementString(string bankStatementString)
        {
            // parse "UTGÅENDE SALDO 20.12.2017"

            Regex regex = new Regex(@".*(\d{2}\.\d{2}\.\d{4})");
            Match match = regex.Match(bankStatementString);
            if (match.Success)
            {
                var dateString = match.Groups[1].Value;
                try
                {
                    return DateTime.ParseExact(dateString, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                    return DateTime.MinValue;
                }
            }
            return DateTime.MinValue;
        }

        public static T GetField<T>(IXLTableRow row, string fieldName)
        {
            object value;
            try
            {
                var item = row.Field(fieldName);
                if (item.HasFormula)
                {
                    value = item.CachedValue;
                }
                else
                {
                    value = item.Value;
                }
            }
            catch (Exception)
            {
                return default(T);
            }

            if (null != value && !"".Equals(value))
            {
                return (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture);
            }
            return default(T);
        }
    }
}
