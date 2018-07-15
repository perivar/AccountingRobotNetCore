using System;
using System.Globalization;

namespace AccountingServices.Oberlo
{
    public class OberloOrder
    {
        public string OrderNumber { get; set; }
        public DateTime CreatedDate { get; set; }
        public string FinancialStatus { get; set; } // 6 = Refunded, 5 = Partially Refunded, 4 = Paid, 
        public string FulfillmentStatus { get; set; } // 1 = Shipped, 2 = In Processing
        public string Supplier { get; set; }
        public string SKU { get; set; }
        public string ProductName { get; set; }
        public string Variant { get; set; }
        public int Quantity { get; set; }
        public decimal ProductPrice { get; set; }
        public string TrackingNumber { get; set; }
        public string AliOrderNumber { get; set; }
        public string CustomerName { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerAddress2 { get; set; }
        public string CustomerCity { get; set; }
        public string CustomerZip { get; set; }
        public string OrderNote { get; set; }
        public string OrderState { get; set; }
        public decimal TotalPrice { get; set; }
        public decimal Cost { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1:yyyy-MM-dd} {2} {3} {4} x {5} {6}", OrderNumber, CreatedDate, AliOrderNumber, SKU, CustomerName, Quantity, Cost.ToString("C", new CultureInfo("en-US")));
        }
    }
}
