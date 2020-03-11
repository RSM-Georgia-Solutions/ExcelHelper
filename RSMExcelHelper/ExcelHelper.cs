using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace RSMExcelHelper
{
    public class Helper
    {
        public void GetAndSave(string sourceLocationPath, string targetLocationPath)
        {
            DateTime now = DateTime.Now;
            Excel.Application xlApp = new Excel.Application { DisplayAlerts = false };
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(sourceLocationPath);
            string fileName = $"{targetLocationPath}-{now.Day - 1}-{now.Month}-{now.Year}.xlsx";
            if (File.Exists(fileName))
            {
                Excel.Workbook xlWorkBook2 = xlApp.Workbooks.Open(fileName);
                if (xlWorkBook2.ReadOnly)
                {
                    xlWorkBook2.Close(true);
                    return;
                }
            }
            xlWorkBook.SaveAs(fileName);
            xlWorkBook.Close(true);
            File.SetAttributes(fileName, FileAttributes.ReadOnly);
        }
        public void RefreshFile(string sourceLocationPath)
        {
            Excel.Application xlApp = new Excel.Application { DisplayAlerts = false };
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(sourceLocationPath);
            xlWorkBook.RefreshAll();
            xlApp.Application.CalculateUntilAsyncQueriesDone();
            xlWorkBook.Save();
            xlWorkBook.Close();
            xlApp.Quit();
        }
    }
}

