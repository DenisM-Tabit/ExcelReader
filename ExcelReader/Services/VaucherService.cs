using ExcelReader.Models;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Text;
using System;
using System.Configuration;
using RestSharp;
using RestSharp.Authenticators;

namespace ExcelReader.Services
{
    internal class VaucherService
    {
        string baseUrl = ConfigurationManager.AppSettings["baseUrl"];
        string testUrl = ConfigurationManager.AppSettings["testUrl"];
        string env = ConfigurationManager.AppSettings["env"];
        string joinChannelGuid = ConfigurationManager.AppSettings["joinChannelGuid"];
        public int totalVouchersSent = 0;
        public int totalVouchersFailedToSend = 0;
        public int totalDuplicateVouchers = 0;

        public void GrantVouchers(List<GrantVoucher> vauchers, BusinessService businessService)
        {
            Console.WriteLine("                       Sending Vouchers To Database ");
            Console.WriteLine("******************************************************************************\n");
            foreach (GrantVoucher voucher in vauchers)
            {
                try
                {
                    if (isDuplicate(voucher))
                    {
                        Console.WriteLine("Voucher already exist in Database, phone number: " + voucher.mobile);
                        continue;
                    }
                    string siteId = businessService.BusinessDetails[voucher.loadBusinessId];
                    bool isVaucherSent = sendVoucher(voucher, siteId, businessService.AccountGuid);
                    Console.ForegroundColor = isVaucherSent ? ConsoleColor.Green : ConsoleColor.Red;
                    Console.WriteLine(isVaucherSent ? "  * Voucher sent successfully" : "  * Failed to send voucher, phone number: " + voucher.mobile);
                }
                catch
                {
                    totalVouchersFailedToSend++;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Failed to send voucher, phone number: " + voucher.mobile);
                    Logger.summuryLog.Add("Failed to send voucher, phone number: " + voucher.mobile);
                }

            }
            LogSummury();

        }
        private void LogSummury()
        {
            Console.ResetColor();
            Console.WriteLine("\n******************************************************************************");
            Console.BackgroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("                               APP FINISHED WORKING!");
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("******************************************************************************\n");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("     number of vouchers sent successfully:   " + totalVouchersSent);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n     number of vouchers failed to send:      " + totalVouchersFailedToSend);
            Console.WriteLine("\n     number of existing vouchers:            " + totalDuplicateVouchers);
            Console.ResetColor();
            Console.WriteLine("\n******************************************************************************");
        }
        private bool isDuplicate(GrantVoucher voucher)
        {
            try
            {
                string sql = @"SELECT top 1 bod.PreDefinedReasonId FROM dbo.BenefitOrderData bod WHERE bod.PreDefinedReasonId = @PreDefinedReasonId";
                SqlCommand ObjCmd = new SqlCommand(sql);
                ObjCmd.Parameters.AddWithValue("PreDefinedReasonId", voucher.preDefinedReasonId);
                var res = Dal.getString(ObjCmd, "preDefinedReasonId");
                if (String.IsNullOrEmpty(res))
                    return false;

                totalDuplicateVouchers++;
                Logger.summuryLog.Add($"   * Already in DB, mobile: {voucher.mobile}");
                return true;
            }
            catch
            {
                return false;
            }
    
        }
        private bool sendVoucher(GrantVoucher voucher, string siteId, string accountGuid)
        {
            string url = baseUrl + "/voucher/grant-voucher";
            string token = "0hvJTEci31Mvze0oVEYCVh7wO4YZnW";
            var jsonVaucher = JsonConvert.SerializeObject(voucher);

            var client = new RestClient(url);
            client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(token, "Bearer");

            var request = new RestRequest(Method.POST);
            request.RequestFormat = DataFormat.Json;
            request.AddJsonBody(jsonVaucher);
            request.AddHeader("siteId", siteId);
            request.AddHeader("joinChannelGuid", joinChannelGuid);
            request.AddHeader("accountGuid", accountGuid);
            request.AddHeader("env", env);

            var response = client.Execute(request);
            if (response.IsSuccessful)
            {
                totalVouchersSent++;
                return true;
            }
            Logger.summuryLog.Add($"   * Failed to send, mobile: {voucher.mobile}");
            totalVouchersFailedToSend++;
            return false;
        }

        public List<GrantVoucher> ConvertExcelObjToApiObj(List<ExcelVoucher> excelVouchers)
        {
            List<GrantVoucher> grantVouchers = new List<GrantVoucher>();

            foreach(ExcelVoucher excelVoucher in excelVouchers)
            {
                var name = getFirstLastName(excelVoucher.CustomerName);
                GrantVoucher grantVoucher = new GrantVoucher()
                {
                    firstName = name.firstName,
                    lastName = name.lastName,
                    mobile = excelVoucher.Mobile.Replace("-", ""),
                    reason = excelVoucher.Reason,
                    issuerName = "",    // can be anything
                    issuerRosUserId = "", // can be anything 
                    loadBusinessId = int.Parse(excelVoucher.BusinessId),
                    preDefinedReasonId = getPreDefinedReasonId(excelVoucher.IssuedDate),
                    preDefinedReasonName = "Granted from import voucher proccess",                  
                    vouchers = new Voucher[1] { new Voucher()
                    {
                        externalSaleId = int.Parse(excelVoucher.ExternalSaleId),
                        benefitMoneyAmount = int.Parse(excelVoucher.MoneyAmount),
                        quantity = 1
                    }},
                    notifyBySms = false
                };
                grantVouchers.Add(grantVoucher);
            }

            return grantVouchers;
        }
        private string getPreDefinedReasonId (DateTime issuedDate)
        {
            return String.Concat(issuedDate.Year, issuedDate.Month, issuedDate.Day, issuedDate.Hour, issuedDate.Minute);
                
        }
        private (string firstName, string lastName) getFirstLastName(string fullName)
        {
            var value = fullName.Split(' ');

            string firstName, lastName; 

            // only first name available 
            if (value.Length == 1)
            {
                firstName = value[0];
                lastName = String.Empty;
            }
            else
            {
                // first name is the last word in the string
                firstName = value[value.Length - 1];
                lastName = String.Join(' ', value, 0, value.Length - 1);    
                // add logic for lastName here 
            }
            return (firstName, lastName);
        }
    }
}
// joinChannelGuid get hard code form app.config  --> done
// accountId get hard code form app.config  --> done 
// get accountGuid by account id from account table -->done
// enviroment get hard code form app.config make it il --> done
// siteId comes with businessId validates that both exist --> done
// TabitOrganizationId from business table --> done