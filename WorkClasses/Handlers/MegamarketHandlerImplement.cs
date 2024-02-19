using BalanceApp.WorkClasses.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class MegamarketHandlerImplement : MarketplaceHandler
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

        public override string GetTitle => "Мегамаркет";

        
        public override void FillBalance()
        {
            if(string.IsNullOrEmpty(PathToBalanceTemplate) || isFill)
            {
                return;
            }

            Excel.Application excelApp = ExcelApplication.GetApp;

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToBalanceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Список товаров"];
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "G"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 3; i <= lastRow; i++)
            {
                int currentBalance = base.GetDataFrom1c(excelSheet, $"G{i}", $"E{i}", true);
                double currentPrice = Math.Ceiling(base.GetDataFrom1c(excelSheet, $"G{i}", $"E{i}", false) * Percent);
                excelSheet.Range[$"J{i}"].Value = currentBalance;
                excelSheet.Range[$"B{i}"].Value = currentBalance > 0 ? "Доступен" : "Не доступен";
                if (currentPrice > 0)
                {
                    excelSheet.Range[$"H{i}"].Value = currentPrice;
                    excelSheet.Range[$"I{i}"].Value = Math.Ceiling(currentPrice * 0.125) * 10;
                }
            }

            string filename = Global.DirectoryToSave 
                + "Остатки ММ " 
                + DateTime.Now.ToShortDateString()
                + PathToBalanceTemplate.Substring(PathToBalanceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
            //чтобы не заполнять шаблон повторно
            isFill = true;
        }

        public override void FillPrice()
        {
            FillBalance();
        }
    }
}
