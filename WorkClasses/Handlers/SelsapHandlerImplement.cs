using BalanceApp.WorkClasses.Models;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class SelsapHandlerImplement : MarketplaceHandler
    {
        public override string GetTitle => "Селсап";
        public double PercentOzon1 { get; set; } = 1.05;
        public double PercentYandex { get; set; } = 1.05;
        public double PercentWildberries { get; set; } = 1.1;
        
        public override void FillBalance()
        {
            if(string.IsNullOrEmpty(PathToBalanceTemplate))
            {
                return; 
            }

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToBalanceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Лист1"];
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 2; i <= lastRow; i++)
            {
                int currentBalance = base.GetDataFrom1c(excelSheet, $"B{i}", $"C{i}", true);
                excelSheet.Range[$"I{i}"].Value = currentBalance;
            }

            string filename = Global.DirectoryToSave 
                + "Остатки селсап " 
                + DateTime.Now.ToShortDateString() 
                + PathToBalanceTemplate.Substring(PathToBalanceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }

        public override void FillPrice()
        {
            if (string.IsNullOrEmpty(PathToPriceTemplate))
            {
                return;
            }

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToPriceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Цены"];
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "EC"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 4; i <= lastRow; i++)
            {
                int currentMinPriceFrom1C = base.GetDataFrom1c(excelSheet, $"EC{i}", $"C{i}", false);

                var currentMinPrice = excelSheet.Range[$"K{i}"].Value;
                if (
                       //если цена есть в 1с и
                       currentMinPriceFrom1C > 0 &&
                       //если текущая ячейки с минимальной ценой пуста
                       (currentMinPrice is null ||
                       //или если она не пуста и ее значение отличается от данных 1с
                       (currentMinPrice is not null && Convert.ToInt32(currentMinPrice) != currentMinPriceFrom1C))
                   )
                {
                    //записывааем минимальную цену
                    excelSheet.Range[$"K{i}"].Value = currentMinPriceFrom1C;
                    //проверяем наличие карточки ВБ
                    if(excelSheet.Range[$"DT{i}"].Value is not null)
                    {
                        excelSheet.Range[$"AN{i}"].Value = Math.Ceiling(currentMinPriceFrom1C * PercentWildberries); 
                    }
                    //проверяем наличие карточки Яндекса
                    if(excelSheet.Range[$"DW{i}"].Value is not null)
                    {
                        double price = Math.Ceiling(currentMinPriceFrom1C * PercentYandex);
                        excelSheet.Range[$"CQ{i}"].Value = price;
                        excelSheet.Range[$"CM{i}"].Value = Math.Ceiling(price * 0.125) * 10;
                    }
                    //проверяем наличие карточки Озон
                    if(excelSheet.Range[$"DU{i}"].Value is not null)
                    {
                        double price = Math.Ceiling(currentMinPriceFrom1C * PercentOzon1);
                        excelSheet.Range[$"BS{i}"].Value = price;
                        excelSheet.Range[$"BO{i}"].Value = Math.Ceiling(price * 0.125) * 10;
                    }

                }
                //если не изменилась, то удаляем эту строку
                //else
                //{
                //    Excel.Range toRemove = excelSheet.get_Range($"A{i}");
                //    toRemove.EntireRow.Delete(Type.Missing);
                //    i--;
                //    lastRow--;
                //}
            }

            string filename = Global.DirectoryToSave 
                + $"Цены селсап " 
                + DateTime.Now.ToShortDateString() 
                + PathToPriceTemplate.Substring(PathToPriceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }
    }
}
