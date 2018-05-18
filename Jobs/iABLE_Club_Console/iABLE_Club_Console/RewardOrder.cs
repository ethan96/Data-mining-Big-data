using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class RewardOrder
    {
        public string orderNo { get; set; }

        public string Receiver { get; set; }

        public string Email { get; set; }

        public string CompanyName { get; set; }

        public string Zip { get; set; }

        public string address { get; set; }

        public string Mobile { get; set; }

        public string Tel { get; set; }

        public string SoMsg { get; set; }

        public string DnMsg { get; set; }

        public string DN { get; set; }

        public string SalesEmail { get; set; }

        public bool SiebelActivityStatus { get; set; }

        public string SiebelActivityMessage { get; set; }

        public string StoreID { get; set; }

        public int RecordID { get; set; }

        public List<RewardRecord> items { get; set; }

        public string SalesDivCode { get; set; }

        public RewardOrder()
        {
            this.items = new List<RewardRecord>();
        }
    }
}
