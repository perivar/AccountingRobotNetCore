using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingServices.PayPalService
{
    public class PayPalTransaction
    {
        public string TransactionID { get; set; }
        public DateTime Timestamp { get; set; }
        public string Status { get; set; }
        public string Type { get; set; }
        public decimal GrossAmount { get; set; }
        public string GrossAmountCurrencyId { get; set; }
        public decimal NetAmount { get; set; }
        public string NetAmountCurrencyId { get; set; }
        public decimal FeeAmount { get; set; }
        public string FeeAmountCurrencyId { get; set; }
        public string Payer { get; set; }
        public string PayerDisplayName { get; set; }

        public override string ToString()
        {
            return string.Format("{0:dd.MM.yyyy} {1} {2} {3:N} {4:N} {5:N} {6}", Timestamp, Type, Status, GrossAmount, NetAmount, FeeAmount, PayerDisplayName);
        }
    }
}
