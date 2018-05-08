using System;
using System.Collections.Generic;
using System.Globalization;
using PayPal.Api;

namespace AccountingServices
{
    public static class Paypal
    {
        public static List<PayPalTransaction> GetPayPalTransactions(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            // get paypal configuration parameters
            string payPalApiUsername = configuration.GetValue("PayPalApiUsername");
            string payPalApiPassword = configuration.GetValue("PayPalApiPassword");
            string payPalApiSignature = configuration.GetValue("PayPalApiSignature");

            var payPalTransactions = new List<PayPalTransaction>();
            return payPalTransactions;
        }
    }
}
