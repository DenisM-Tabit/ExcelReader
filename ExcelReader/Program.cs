using System;
using System.Collections.Generic;
using System.Configuration;
using ExcelDataReader;
using ExcelReader.Models;
using ExcelReader.Services;

namespace ExcelReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string accountId =  ConfigurationManager.AppSettings["accountId"];
            Console.WriteLine("******************************************************************************");
            Console.WriteLine("                          Excel Reader Started");
            Console.WriteLine("******************************************************************************\n");

            VaucherService voucherService = new VaucherService();
            BusinessService businessService = new BusinessService(accountId);    


            //get the list of vouchers from excel
            
            FileReader reader = new FileReader(businessService);
            List<ExcelVoucher> vauchers = reader.ReadFile();

            Console.WriteLine("  Vouchers loaded: " + vauchers.Count);
            Console.WriteLine("  Vouchers failed to load: " + Logger.badDataLog.Count);
            Console.WriteLine("******************************************************************************\n");

            // convert the excel object to api objects 
            List<GrantVoucher> grantVouchersList = voucherService.ConvertExcelObjToApiObj(vauchers);

            Console.ForegroundColor = ConsoleColor.Red;
            Logger.LogFaultyDataInExcel();
            Console.ResetColor();
            // send vouchers 
            voucherService.GrantVouchers(grantVouchersList, businessService);
            Console.ResetColor();
            Logger.CreateTxtLog(voucherService.totalVouchersSent, voucherService.totalVouchersFailedToSend, voucherService.totalDuplicateVouchers);
            Console.ReadKey();


        }
    }
}
