using Stripe;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AccountingServices
{
    public static class Stripe
    {
        #region Charge Transactions (StripeChargeService)
        public static List<StripeTransaction> GetStripeChargeTransactions(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            // get stripe configuration parameters
            string stripeApiKey = configuration.GetValue("StripeApiKey");

            StripeConfiguration.SetApiKey(stripeApiKey);

            var chargeService = new StripeChargeService();
            chargeService.ExpandBalanceTransaction = true;
            chargeService.ExpandCustomer = true;
            chargeService.ExpandInvoice = true;

            var allCharges = new List<StripeCharge>();
            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeCharge> charges = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    charges = chargeService.List(
                    new StripeChargeListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        Created = new StripeDateFilter
                        {
                            GreaterThanOrEqual = from,
                            LessThanOrEqual = to
                        }
                    });
                    itemsExpected = charges.Count();
                }
                else
                {
                    charges = chargeService.List(
                    new StripeChargeListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        StartingAfter = lastId,
                        Created = new StripeDateFilter
                        {
                            GreaterThanOrEqual = from,
                            LessThanOrEqual = to
                        }
                    });
                    itemsExpected = charges.Count();
                }

                allCharges.AddRange(charges);
                if (allCharges.Count() > 0) lastId = charges.LastOrDefault().Id;
            }

            var stripeTransactions = new List<StripeTransaction>();
            foreach (var charge in allCharges)
            {
                // only process the charges that were successfull, aka Paid
                if (charge.Paid)
                {
                    var stripeTransaction = new StripeTransaction();
                    stripeTransaction.TransactionID = charge.Id;
                    stripeTransaction.Created = charge.Created;
                    stripeTransaction.Paid = charge.Paid;
                    stripeTransaction.CustomerEmail = charge.Metadata["email"];
                    stripeTransaction.OrderID = charge.Metadata["order_id"];
                    stripeTransaction.Amount = (decimal)charge.Amount / 100;
                    stripeTransaction.Net = (decimal)charge.BalanceTransaction.Net / 100;
                    stripeTransaction.Fee = (decimal)charge.BalanceTransaction.Fee / 100;
                    stripeTransaction.Currency = charge.Currency;
                    stripeTransaction.Description = charge.Description;
                    stripeTransaction.Status = charge.Status;
                    decimal amountRefunded = (decimal)charge.AmountRefunded / 100;
                    if (amountRefunded > 0)
                    {
                        // if anything has been refunded
                        // don't add if amount refunded and amount is the same (full refund)
                        if (stripeTransaction.Amount - amountRefunded == 0)
                        {
                            continue;
                        }
                        else
                        {
                            stripeTransaction.Amount = stripeTransaction.Amount - amountRefunded;
                            stripeTransaction.Net = stripeTransaction.Net - amountRefunded;
                        }
                    }
                    stripeTransactions.Add(stripeTransaction);
                }
            }
            return stripeTransactions;
        }
        #endregion

        #region Payout Transactions (StripeBalanceService)
        public static List<StripeTransaction> GetStripePayoutTransactions(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            // get stripe configuration parameters
            string stripeApiKey = configuration.GetValue("StripeApiKey");

            StripeConfiguration.SetApiKey(stripeApiKey);

            var balanceService = new StripeBalanceService();
            var allBalanceTransactions = new List<StripeBalanceTransaction>();
            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeBalanceTransaction> charges = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    charges = balanceService.List(
                    new StripeBalanceTransactionListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        Created = new StripeDateFilter
                        {
                            GreaterThanOrEqual = from,
                            LessThanOrEqual = to
                        }
                    });
                    itemsExpected = charges.Count();
                }
                else
                {
                    charges = balanceService.List(
                    new StripeBalanceTransactionListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        StartingAfter = lastId,
                        Created = new StripeDateFilter
                        {
                            GreaterThanOrEqual = from,
                            LessThanOrEqual = to
                        }
                    });
                    itemsExpected = charges.Count();
                }

                allBalanceTransactions.AddRange(charges);
                if (allBalanceTransactions.Count() > 0) lastId = charges.LastOrDefault().Id;
            }

            var stripeTransactions = new List<StripeTransaction>();
            foreach (var balanceTransaction in allBalanceTransactions)
            {
                // only process the charges that are payouts
                if (balanceTransaction.Type == "payout")
                {
                    var stripeTransaction = new StripeTransaction();
                    stripeTransaction.TransactionID = balanceTransaction.Id;
                    stripeTransaction.Created = balanceTransaction.Created;
                    stripeTransaction.AvailableOn = balanceTransaction.AvailableOn;
                    stripeTransaction.Paid = (balanceTransaction.Status == "available");
                    stripeTransaction.Amount = (decimal)balanceTransaction.Amount / 100;
                    stripeTransaction.Net = (decimal)balanceTransaction.Net / 100;
                    stripeTransaction.Fee = (decimal)balanceTransaction.Fee / 100;
                    stripeTransaction.Currency = balanceTransaction.Currency;
                    stripeTransaction.Description = balanceTransaction.Description;
                    stripeTransaction.Status = balanceTransaction.Status;

                    stripeTransactions.Add(stripeTransaction);
                }
            }
            return stripeTransactions;
        }
        #endregion
    }
}
