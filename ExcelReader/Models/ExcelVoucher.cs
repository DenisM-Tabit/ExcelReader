using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    internal class ExcelVoucher
    {
        public string Mobile { get; set; }
        public string CustomerName { get; set; }
        public DateTime IssuedDate { get; set; }
        public string ExternalSaleId { get; set; }
        public string BusinessId { get; set; }
        public string MoneyAmount { get; set; }
        public string Reason { get; set; }


    }
}
