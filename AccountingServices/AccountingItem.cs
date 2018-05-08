using System;
using CsvHelper.Configuration;

namespace AccountingServices
{
    public class AccountingItem
    {
        public int Periode => Date.Month;
        public DateTime Date { get; set; }

        public int Number { get; set; }
        public long ArchiveReference { get; set; }
        public string TransactionID { get; set; }
        public string Type { get; set; } // Overføring (intern), Overførsel (ekstern), Visa, Avgift
        public string AccountingType { get; set; }
        public string Text { get; set; }
        public string CustomerName { get; set; }
        public string ErrorMessage { get; set; }

        public string Gateway { get; set; }

        public string NumSale { get; set; }
        public string NumPurchase { get; set; }

        public decimal PurchaseOtherCurrency { get; set; }
        public string OtherCurrency { get; set; }

        public decimal AccountPaypal { get; set; }              // 1910
        public decimal AccountStripe { get; set; }              // 1912
        public decimal AccountVipps { get; set; }               // 1914
        public decimal AccountBank { get; set; }                // 1920

        public decimal VATPurchase { get; set; }
        public decimal VATSales { get; set; }

        public decimal SalesVAT { get; set; }                   // 3000
        public decimal SalesVATExempt { get; set; }             // 3100

        public decimal CostOfGoods { get; set; }                // 4005
        public decimal CostForReselling { get; set; }           // 4300
        public decimal CostForSalary { get; set; }              // 5000
        public decimal CostForSalaryTax { get; set; }           // 5400
        public decimal CostForDepreciation { get; set; }        // 6000
        public decimal CostForShipping { get; set; }            // 6100
        public decimal CostForElectricity { get; set; }         // 6340 
        public decimal CostForToolsInventory { get; set; }      // 6500 (includes purchasing cell phones etc.)
        public decimal CostForMaintenance { get; set; }         // 6695
        public decimal CostForFacilities { get; set; }          // 6800 

        public decimal CostOfData { get; set; }                 // 6810 
        public decimal CostOfPhoneInternetUse { get; set; }     // 6900
        public decimal PrivateUseOfECom { get; set; }           // 7098 (tilbakeføringskonto vedr. EKOM)
        public decimal CostForTravelAndAllowance { get; set; }  // 7140
        public decimal CostOfAdvertising { get; set; }          // 7330
        public decimal CostOfOther { get; set; }                // 7700

        public decimal FeesBank { get; set; }                   // 7770
        public decimal FeesPaypal { get; set; }                 // 7780
        public decimal FeesStripe { get; set; }                 // 7785 

        public decimal CostForEstablishment { get; set; }       // 7790

        public decimal IncomeFinance { get; set; }              // 8099
        public decimal CostOfFinance { get; set; }              // 8199

        public decimal Investments { get; set; }                // 1200
        public decimal AccountsReceivable { get; set; }         // 1500
        public decimal PersonalWithdrawal { get; set; }
        public decimal PersonalDeposit { get; set; }

        public bool Equals(AccountingItem other)
        {
            if (other == null) return false;

            return
                ArchiveReference == other.ArchiveReference &&
                //TransactionID == other.TransactionID &&
                //Type == other.Type &&
                Date == other.Date &&
                string.Equals(Text, other.Text) &&
                                                //AccountPaypal == other.AccountPaypal &&
                                                //AccountStripe == other.AccountStripe &&
                                                //AccountVipps == other.AccountVipps &&
                                                AccountBank == other.AccountBank;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals(obj as AccountingItem);
        }

        public override int GetHashCode()
        {
            // http://www.aaronstannard.com/overriding-equality-in-dotnet/
            // https://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode/263416#263416
            unchecked
            {
                var hashCode = 13;
                hashCode = (hashCode * 397) ^ ArchiveReference.GetHashCode();
                //hashCode = (hashCode * 397) ^ (!string.IsNullOrEmpty(TransactionID) ? TransactionID.GetHashCode() : 0);                
                //hashCode = (hashCode * 397) ^ (!string.IsNullOrEmpty(Type) ? Type.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (!string.IsNullOrEmpty(Text) ? Text.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ Date.GetHashCode();
                //hashCode = (hashCode * 397) ^ AccountPaypal.GetHashCode();
                //hashCode = (hashCode * 397) ^ AccountStripe.GetHashCode();
                //hashCode = (hashCode * 397) ^ AccountVipps.GetHashCode();
                hashCode = (hashCode * 397) ^ AccountBank.GetHashCode();

                return hashCode;
            }
        }

        public override string ToString()
        {
            return string.Format("{0} {1:dd.MM.yyyy} {2} {3:C} {4:C} {5:C} {6:C}", Number, Date, Text, AccountPaypal, AccountStripe, AccountVipps, AccountBank);
        }
    }

    public sealed class AccountingItemCsvMap : ClassMap<AccountingItem>
    {
        public AccountingItemCsvMap()
        {
            Map(m => m.Periode);
            Map(m => m.Date).TypeConverterOption.Format("yyyy.MM.dd");

            Map(m => m.Number);
            Map(m => m.ArchiveReference);
            Map(m => m.Type);
            Map(m => m.AccountingType);
            Map(m => m.Text);
            Map(m => m.CustomerName);
            Map(m => m.ErrorMessage);

            Map(m => m.Gateway);
            Map(m => m.NumSale);
            Map(m => m.NumPurchase);
            Map(m => m.PurchaseOtherCurrency);
            Map(m => m.OtherCurrency);

            Map(m => m.AccountPaypal);
            Map(m => m.AccountStripe);
            Map(m => m.AccountVipps);
            Map(m => m.AccountBank);

            Map(m => m.VATPurchase);
            Map(m => m.VATSales);

            Map(m => m.SalesVAT);
            Map(m => m.SalesVATExempt);

            Map(m => m.CostOfGoods);
            Map(m => m.CostForReselling);
            Map(m => m.CostForSalary);
            Map(m => m.CostForSalaryTax);
            Map(m => m.CostForDepreciation);
            Map(m => m.CostForShipping);
            Map(m => m.CostForElectricity);
            Map(m => m.CostForToolsInventory);
            Map(m => m.CostForMaintenance);
            Map(m => m.CostForFacilities);

            Map(m => m.CostOfData);
            Map(m => m.CostOfPhoneInternetUse);
            Map(m => m.PrivateUseOfECom);
            Map(m => m.CostForTravelAndAllowance);
            Map(m => m.CostOfAdvertising);
            Map(m => m.CostOfOther);

            Map(m => m.FeesBank);
            Map(m => m.FeesPaypal);
            Map(m => m.FeesStripe);

            Map(m => m.CostForEstablishment);

            Map(m => m.IncomeFinance);
            Map(m => m.CostOfFinance);

            Map(m => m.Investments);
            Map(m => m.AccountsReceivable);
            Map(m => m.PersonalWithdrawal);
            Map(m => m.PersonalDeposit);
        }
    }
}
