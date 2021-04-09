using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;

namespace DeactivateUser
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;


            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Project\CRM2015\CRM2016\Utilities\Arup.DeactivateUser\DeactivateUser\DeactivateUser\Excel\training to prod user.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);

            range = xlWorkSheet.UsedRange;
            var rw = range.Rows.Count;
            var cl = range.Columns.Count;


            for (var rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (var cCnt = 1; cCnt <= cl; cCnt++)
                {
                    var str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    Console.WriteLine(rCnt + " " + str);
                    
                }
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            Console.Read();

        }
    }
}
