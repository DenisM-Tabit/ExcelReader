using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Services
{ 
    internal class BusinessService
    {
        public string AccountId;
        public Dictionary<int, string> BusinessDetails;
        public List<int> ExternalId;
        public string AccountGuid;
        public string FilePath;
        public BusinessService(string accountId)
        {
            AccountId = accountId;
            BusinessDetails = GetAllBusiness();
            ExternalId = GetAllExternalId();
            AccountGuid = GetAccountGuid();
            FilePath = String.Concat(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["excelFileName"]) ;
        }
        public bool isVoucherUploaded()
        {
            return false;
        }
        private Dictionary<int, string> GetAllBusiness()
        {
            string sql = @"SELECT businessId, TabitOrganizationId as siteId FROM Business b WHERE b.AccountId = @AccountId";
            SqlCommand ObjCmd = new SqlCommand(sql);
            ObjCmd.Parameters.AddWithValue("AccountId", int.Parse(AccountId));

            return Dal.getIntDictionary(ObjCmd, "businessId");
        }

        private List<int> GetAllExternalId()
        {
            string sql = @"SELECT externalSaleId FROM externalSale WHERE accountId = @AccountId";
            SqlCommand ObjCmd = new SqlCommand(sql);
            ObjCmd.Parameters.AddWithValue("AccountId", AccountId);

            return Dal.getIntList(ObjCmd, "externalSaleId");
        }

        private string GetAccountGuid()
        {
            string sql = @"SELECT accountGuid from Account where accountId = @AccountId";
            SqlCommand ObjCmd = new SqlCommand(sql);
            ObjCmd.Parameters.AddWithValue("AccountId", AccountId);

            return Dal.getString(ObjCmd, "accountGuid");
        }
    }
}
