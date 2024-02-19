using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses
{
    public static class ExcelApplication
    {
        static Excel.Application excelApp = null!;
        static Process excelProc = null!;
        public static Excel.Application GetApp
        {
            get
            {
                if (excelApp is null)
                {
                    excelApp = new Excel.Application();
                    excelProc = Process.GetProcessesByName("EXCEL").Last();
                }
                return excelApp;
            }
        }

        public static void Close()
        {
            if (excelApp != null)
            {
                excelApp.Application.Quit();
                excelProc.Kill();
            }
        }
    }
}
