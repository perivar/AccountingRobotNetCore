using System;

namespace AccountingServices.ShopifyService
{
    public class ShopifyOrder
    {
        public long Id { get; set; }

        public DateTime CreatedAt { get; set; }
        public DateTime ProcessedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
        public DateTime CancelledAt { get; set; }

        public string Name { get; set; }
        public string FinancialStatus { get; set; }
        public string FulfillmentStatus { get; set; }

        public string Gateway { get; set; }
        public string PaymentId { get; set; }

        public decimal TotalPrice { get; set; }
        public decimal TotalTax { get; set; }

        public string CustomerEmail { get; set; }
        public string CustomerName { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerAddress2 { get; set; }
        public string CustomerZipCode { get; set; }
        public string CustomerCity { get; set; }

        public string Note { get; set; }

        public bool Cancelled
        {
            get
            {
                return (CancelledAt != default(DateTime) ? true : false);
            }
        }

        public override string ToString()
        {
            return string.Format("{0} {1:dd-MM-yyyy} {2} {3} {4} {5} {6:C} {7}", Id, CreatedAt, Name, FinancialStatus, FulfillmentStatus, Gateway, TotalPrice, CustomerName);
        }
    }
}
