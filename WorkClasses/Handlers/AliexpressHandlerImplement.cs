using BalanceApp.WorkClasses.Models;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class AliexpressHandlerImplement : MarketplaceHandler
    {
        bool isFill;
        public new string PathToPriceTemplate
        {
            get
            {
                return PathToBalanceTemplate;
            }
            set
            {
                PathToBalanceTemplate = value;
            }
        }

        public override string GetTitle => "Алиэкспресс";
        public override void FillBalance()
        {
            if(!string.IsNullOrEmpty(PathToBalanceTemplate) && !isFill)
            {
                Excel.Application excelApp = ExcelApplication.GetApp;

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToBalanceTemplate);
                Excel.Worksheet excelSheet = excelWorkbook.ActiveSheet;
                int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "E"].End[Excel.XlDirection.xlUp].Row;

                for (int i = 4; i <= lastRow; i++)
                {
                    int currentBalance = base.GetDataFrom1c(excelSheet, $"E{i}", $"C{i}", true);
                    double currentPrice = Math.Ceiling(base.GetDataFrom1c(excelSheet, $"E{i}", $"C{i}", false) * Percent);

                    excelSheet.Range[$"H{i}"].Value = currentBalance;
                    if (currentPrice > 0)
                    {
                        excelSheet.Range[$"J{i}"].Value = currentPrice;
                        excelSheet.Range[$"I{i}"].Value = Math.Ceiling(currentPrice * 0.125) * 10;
                    }
                }

                string filename = Global.DirectoryToSave 
                    + "Остатки Али " 
                    + DateTime.Now.ToShortDateString()
                + PathToBalanceTemplate.Substring(PathToBalanceTemplate.LastIndexOf('.'));
                CheckExists(filename);
                excelApp.Application.ActiveWorkbook.SaveAs(filename);

                excelWorkbook.Close();
                //чтобы не заполнять шаблон повторно
                isFill = true;
            }
        }

        public override void FillPrice()
        {
            FillBalance();
        }
    }
}
