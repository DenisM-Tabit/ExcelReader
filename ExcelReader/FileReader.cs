using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Office.Interop.Excel;
using ExcelReader.Models;
using System.Reflection;
using ExcelReader.Services;

namespace ExcelReader
{
    internal class FileReader
    {
        Application excel = new Application();
        List<string> excelTitleList = new List<string>();
        List<ExcelVoucher> vouchers = new List<ExcelVoucher>();
        Workbook wb;
        Worksheet ws;
        BusinessService businessServiceObject;

        public FileReader(BusinessService businessService)
        {
            businessServiceObject = businessService;
            wb = excel.Workbooks.Open(businessServiceObject.FilePath);
            ws = wb.Worksheets[1];
        }

        public List<ExcelVoucher> ReadFile()
        {
            populateTitleList();
            int columnsInExcel = excelTitleList.Count;

            // starting from row 2 as row 1 is the titles
            int row = 2;
            int emptyRowsCount = 0;
            while (emptyRowsCount <= 5)
            {
                if(ws.Cells[row, 1].Value == null)
                    emptyRowsCount++;
                else
                {
                    emptyRowsCount = 0;
                    var voucherObj = getVoucherFromExcel(columnsInExcel, row);
                    ExcelVoucher compensation = voucherObj.voucher;
                    populateCompensationList(row, compensation, voucherObj.isValidData);
                }
                row++;
            }
            CloseExcel(excel);
            return vouchers;
        }
        private (ExcelVoucher voucher, bool isValidData) getVoucherFromExcel(int columnsInExcel, int row)
        {
            ExcelVoucher voucher = new ExcelVoucher();
            bool isValidData = true;
            for (int col = 1; col <= columnsInExcel; col++)
            {
                try
                {
                    // find the string value of the propery in excelTitleList
                    PropertyInfo columnTitle = typeof(ExcelVoucher).GetProperty(excelTitleList[col - 1]);

                    // validates whatever there is a value and if it's number;
                    if (ws.Cells[row, col].Value == null) throw new Exception();
                    bool isDouble = ws.Cells[row, col].Value is double;
                    

                    // sets the value of compensation object by the columnTitle, if value is number converts it to string
                    columnTitle.SetValue(
                        voucher,
                        isDouble ? ws.Cells[row, col].Value + "" : ws.Cells[row, col].Value
                    );
                }
                catch {
                    isValidData = false;
                }

            }
            return (voucher, isValidData);
        }

        private void populateCompensationList(int row, ExcelVoucher compensation, bool isValidData)
        {
            if (!isValidData)
            {
                Logger.addBadRow("   * Row: " + row + " Bad Data");
                return;
            }
            if (!businessServiceObject.BusinessDetails.ContainsKey(Int32.Parse(compensation.BusinessId)))
                Logger.addBadRow("   * Row: " + row + " Bad BusinessId: " + compensation.BusinessId);
            else if (!businessServiceObject.ExternalId.Contains(Int32.Parse(compensation.ExternalSaleId)))
                Logger.addBadRow("   * Row: " + row + " Bad ExternalSaleId: " + compensation.ExternalSaleId);
            else
                vouchers.Add(compensation);
        }

        private bool validateBusinessId()
        {
            return true;
        }

        private void populateTitleList()
        {
            int currentColumn = 0;
            while (ws.Cells[1 , currentColumn + 1].Value != null)
            {
                excelTitleList.Add(ws.Cells[1, currentColumn + 1].Value);
                currentColumn++;
            }
        }

        private void CloseExcel(Application ExcelApplication = null)
        {
            if (ExcelApplication != null)
            {
                ExcelApplication.Workbooks.Close();
                ExcelApplication.Quit();
            }

            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0) { PK.Kill(); }
            }
        }

    }
}
