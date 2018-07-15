using AccountingServices.Helpers;
using Stripe;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AccountingServices.Stripe
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

            var payoutService = new StripePayoutService();
            var allPayoutTransactions = new List<StripePayout>();
            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripePayout> charges = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    charges = payoutService.List(
                    new StripePayoutListOptions()
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
                    charges = payoutService.List(
                    new StripePayoutListOptions()
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

                allPayoutTransactions.AddRange(charges);
                if (allPayoutTransactions.Count() > 0) lastId = charges.LastOrDefault().Id;
            }

            var stripeTransactions = new List<StripeTransaction>();
            foreach (var payoutTransaction in allPayoutTransactions)
            {
                // only process the charges that are payouts
                if (payoutTransaction.Object == "payout")
                {
                    var stripeTransaction = new StripeTransaction();
                    stripeTransaction.TransactionID = payoutTransaction.Id;
                    stripeTransaction.Created = payoutTransaction.Created;
                    stripeTransaction.AvailableOn = payoutTransaction.ArrivalDate;
                    stripeTransaction.Paid = (payoutTransaction.Status == "paid");
                    stripeTransaction.Amount = (decimal)payoutTransaction.Amount / 100;
                    stripeTransaction.Currency = payoutTransaction.Currency;
                    stripeTransaction.Description = payoutTransaction.Description;
                    stripeTransaction.Status = payoutTransaction.Status;

                    stripeTransactions.Add(stripeTransaction);
                }
            }
            return stripeTransactions;
        }
        #endregion
    }
}
