using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    internal class GrantVoucher
    {
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string mobile { get; set; }
        public string reason { get; set; }
        public string preDefinedReasonId { get; set; }
        public string preDefinedReasonName { get; set; }
        public string issuerName { get; set; }
        public string issuerRosUserId { get; set; }
        public int externalSaleId { get; set; }
        public int benefitMoneyAmount { get; set; }
        public Voucher[] vouchers { get; set; }
        public string orderId { get; set; }
        public int loadBusinessId { get; set; }
        public bool notifyBySms { get; set; }

    }

    internal class Voucher
    {
        public int externalSaleId { get; set; }
        public string name { get; set; }
        public int benefitMoneyAmount { get; set; }
        public int quantity { get; set; }
    }

}
