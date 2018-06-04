using System;
using System.Data;
using System.Data.Common;
using System.Globalization;

namespace AccountingServices
{
    public static class GoogleSheetsUtils
    {
        public static T GetField<T>(DataRow row, string fieldName)
        {
            object value;
            try
            {
                value = row[fieldName];
            }
            catch (Exception)
            {
                return default(T);
            }

            if (value != DBNull.Value && value != null && !"".Equals(value))
            {
                return (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture);
            }
            return default(T);
        }

    }
}