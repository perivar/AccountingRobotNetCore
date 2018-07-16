using System;

namespace AccountingServices.StripeService
{
    public class StripeTransaction
    {
        public string TransactionID { get; set; }
        public string OrderID { get; set; }
        public DateTime Created { get; set; }
        public DateTime AvailableOn { get; set; }
        public bool Paid { get; set; }
        public string CustomerEmail { get; set; }
        public decimal Amount { get; set; }
        public decimal Net { get; set; }
        public decimal Fee { get; set; }
        public string Currency { get; set; }
        public string Description { get; set; }
        public string Status { get; set; }

        public override string ToString()
        {
            return string.Format("{0:dd.MM.yyyy} {1:dd.MM.yyyy} {2} {3:N} {4:N} {5:N} {6} {7}", Created, AvailableOn, Status, Amount, Net, Fee, Description, CustomerEmail);
        }

    }
}
