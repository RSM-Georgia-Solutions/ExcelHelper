using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace RSMExcelHelper
{
    public class Helper
    {
        public void GenerateReadOnlyFile(string sourceLocationPath, string targetLocationPath)
        {
            DateTime now = DateTime.Now;
            Excel.Application xlApp = new Excel.Application { DisplayAlerts = false };
            Excel.Workbooks xlWorkbooks = xlApp.Workbooks;
            Excel.Workbook xlWorkBook = xlWorkbooks.Open(sourceLocationPath);

            string fileName = $"{targetLocationPath}-{now.Day - 1}-{now.Month}-{now.Year}.xlsx";
            if (File.Exists(fileName))
            {
                Excel.Workbook xlWorkBook2 = xlWorkbooks.Open(fileName);
                if (xlWorkBook2.ReadOnly)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlWorkbooks);
                    Marshal.ReleaseComObject(xlApp);
                    return;
                }
            }
            xlWorkBook.SaveAs(fileName);
            xlWorkBook.Close();
            File.SetAttributes(fileName, FileAttributes.ReadOnly);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlApp);
        }
        public void RefreshFile(string sourceLocationPath)
        {
            Excel.Application xlApp = new Excel.Application { DisplayAlerts = false };
            Excel.Workbooks xlWorkbooks = xlApp.Workbooks;
            Excel.Workbook xlWorkBook = xlWorkbooks.Open(sourceLocationPath);
            xlWorkBook.RefreshAll();
            xlApp.Application.CalculateUntilAsyncQueriesDone();
            xlWorkBook.Save();
            xlWorkBook.Close();
            xlWorkbooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlApp);

        }
    }
}

