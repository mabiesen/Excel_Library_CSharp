using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class ExcelService
    {
        public ExcelWorkbook activeWorkbook = new ExcelWorkbook();

        public ExcelService(string filePath)
        {
            //Create App instance, open workbook within app
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            //Get the attributes of workbook.
            //Note: Since Workbooks contain sheets, the workbook code contains sheet creation code.
            //Cycles through each sheet, creates, adds to workbook sheet list
            GetTheWorkbookData(xlWorkBook);

            //Close and release com objects.  
            //Note: worksheet released in workbook section
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }

        public void GetTheWorkbookData(Microsoft.Office.Interop.Excel.Workbook xlWorkBook)
        {
            this.activeWorkbook.workbookName = xlWorkBook.Name;
            this.activeWorkbook.workbookPath = xlWorkBook.Path;

            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

            int worksheetCount;

            worksheetCount = xlWorkBook.Worksheets.Count;
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //collections start at base 1 in excel.  iterate and call function with worksheets
            for (int i = 1; i <= worksheetCount; i++)
            {
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                GetTheWorksheet(xlWorkSheet);
            }
            Marshal.ReleaseComObject(xlWorkSheet);
        }

        public void GetTheWorksheet(Microsoft.Office.Interop.Excel.Worksheet thisWorksheet)
        {
            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            Microsoft.Office.Interop.Excel.Range range;

            range = thisWorksheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            string[,] data = new string[rw, cl];

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    data[rCnt - 1, cCnt - 1] = str;
                }
            }
            //Now create a new worksheet object and insert the worksheet data and name
            ExcelWorksheet newWorksheet = new ExcelWorksheet();
            newWorksheet.worksheetData = data;
            newWorksheet.worksheetName = thisWorksheet.Name;
            newWorksheet.CreateTableFromData(true);
            this.activeWorkbook.worksheetList.Add(newWorksheet);

        }
    }
}
