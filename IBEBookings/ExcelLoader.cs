using System;
using Excel = Microsoft.Office.Interop.Excel;

//namespace TuicContentLoader
//{
public class ExcelLoader
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;

        public Excel.Workbook getWorkbook(string sFile)
        {

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "t", false, false, 0, true, 1, 0);

            return xlWorkBook;
        }

    public bool saveWorkbook() {
        try
        {
            xlWorkBook.Save();
            return true;
        }
        catch (Exception ee)
        {
            Console.WriteLine(ee.Message);
            return false;
        }
    }

        public void quit()
        {
            xlWorkBook.Close(false, null, null);
            xlApp.Quit();
        }
    }
//}
