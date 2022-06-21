using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    internal class Logger
    {
        public static List<string> badDataLog = new List<string>();
        public static List<string> summuryLog = new List<string>();

        public static void addBadRow(string info)
        {
            badDataLog.Add(info);
        }
        public static void LogFaultyDataInExcel()
        {
            Console.WriteLine("  List of faulty data rows in excel");
            Console.WriteLine("******************************************************************************\n");
            foreach (var log in badDataLog)
            {
                Console.WriteLine(log);
            }
            Console.WriteLine("\n******************************************************************************\n\n\n");
        }
        public static void CreateTxtLog(int sentCount, int failedCount, int duplicateCount)
        {
            string path = String.Concat(AppDomain.CurrentDomain.BaseDirectory, "Summury.txt");
            // Create a file to write to.
            using (StreamWriter sw = File.CreateText(path))
            {
                sw.WriteLine("App Finished Working!\n");
                sw.WriteLine("Total Vouchers sent: " + sentCount);
                sw.WriteLine("\nFailed to send: " + failedCount);
                sw.WriteLine("\nAlready exist in DB: " + duplicateCount + "\n");
                sw.WriteLine("\n List of Errors\n*****************************");
                foreach (var log in badDataLog)
                    sw.WriteLine(log);
                foreach (var log in summuryLog)
                    sw.WriteLine(log);
            }
        }
    }
}
