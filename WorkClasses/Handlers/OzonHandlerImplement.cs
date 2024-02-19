using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using BalanceApp.WorkClasses.Models;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class OzonHandlerImplement : MarketplaceHandler
    {
        readonly int shopNumber;
        string pathToBalanceTemplate;
        public new string PathToBalanceTemplate
        {
            set
            {
                try
                {
                    File.Replace(value, pathToBalanceTemplate, value + ".bac");
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            get => pathToBalanceTemplate;
        }

        public delegate void ChangePercent(Marketplace marketplace, double percent);
        public event ChangePercent? OnChangePercent;
        public new double Percent 
        {
            set
            {
                base.Percent = value;
                if (shopNumber == 1)
                    OnChangePercent?.Invoke(Marketplace.Ozon1, value);
            }
            get => base.Percent;
        }

        public string PathToListProducts { get; set; } = null!;

        public override string GetTitle => $"Озон{shopNumber}";

        public OzonHandlerImplement(int shopNumber) 
        {
            this.shopNumber = shopNumber;
            pathToBalanceTemplate = Directory.GetCurrentDirectory() + $@"\Files\Шаблон остатков озон{shopNumber}.xlsx";
        }
        public override void FillBalance()
        {
            if (String.IsNullOrEmpty(PathToBalanceTemplate) || String.IsNullOrEmpty(PathToListProducts))
                return;

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToBalanceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Остатки на складе"];

            Excel.Application excelAppList = new Excel.Application();
            excelAppList.Workbooks.OpenText(PathToListProducts, DataType: XlTextParsingType.xlDelimited, Semicolon: true, Origin: 65001);
            Process excelProc = Process.GetProcessesByName("EXCEL").Last();
            Excel.Worksheet excelSheetList = excelAppList.ActiveWorkbook.ActiveSheet;

            int lastRowList = excelSheetList.Cells[excelSheetList.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 2; i <= lastRowList; i++)
            {
                int currentBalance = base.GetDataFrom1c(excelSheetList, $"F{i}", $"A{i}", true);

                Excel.Range storage = excelWorkbook.Worksheets["Инструкция"].Range[excelSheet.Range[$"A{2}"].Validation.Formula1.Substring(excelSheet.Range[$"A{2}"].Validation.Formula1.IndexOf("!") + 1)];
                for (int s = 1; s <= storage.Count; s++)
                {
                    //Заполняем название склада в шаблоне
                    excelSheet.Range[$"A{(i - 2) * storage.Count + (s - 1) + 2}"].Value = storage[s].Value;
                    //заполняем столбец с id в шаблоне
                    excelSheet.Range[$"B{(i - 2) * storage.Count + (s - 1) + 2}"].Value = excelSheetList.Range[$"A{i}"].Value;
                    //заполняем столбец наименования в шаблоне
                    excelSheet.Range[$"C{(i - 2) * storage.Count + (s - 1) + 2}"].Value = excelSheetList.Range[$"F{i}"].Value;
                    //заполняем столбец с количеством в шаблоне
                    excelSheet.Range[$"D{(i - 2) * storage.Count + (s - 1) + 2}"].Value = currentBalance;

                }
            }
            excelAppList.Application.Quit();
            excelProc.Kill();

            string filename = Global.DirectoryToSave 
                + $"Остатки озон{shopNumber} " 
                + DateTime.Now.ToShortDateString()
                + PathToBalanceTemplate.Substring(PathToBalanceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);
            
            excelWorkbook.Close();
        }
        public override void FillPrice()
        {
            if (string.IsNullOrEmpty(PathToPriceTemplate))
                return;

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToPriceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Товары и цены"];
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "C"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 5; i <= lastRow; i++)
            {
                int currentMinPriceFrom1C = base.GetDataFrom1c(excelSheet, $"C{i}", $"A{i}", false);

                //сравниваем цену с ценой шаблона
                //если она изменилась, то записываем новые данные
                var oldMinPrice = excelSheet.Range[$"Q{i}"].Value;
                if (
                       //если цена есть в 1с и
                       currentMinPriceFrom1C > 0 &&
                       //если текущая ячейки с минимальной ценой пуста
                       (oldMinPrice is null ||
                       //или если она не пуста и ее значение отличается от данных 1с
                       (oldMinPrice is not null && Convert.ToInt32(oldMinPrice) != Math.Ceiling(currentMinPriceFrom1C * Percent)))
                   )
                {
                    double currentMinPrice = Math.Ceiling(currentMinPriceFrom1C * Percent);
                    excelSheet.Range[$"AS{i}"].Value = currentMinPrice.ToString();
                    excelSheet.Range[$"AQ{i}"].Value = excelSheet.Range[$"AQ{i}"].Validation.Formula1.Split(";")[2];
                    excelSheet.Range[$"AT{i}"].Value = excelSheet.Range[$"AT{i}"].Validation.Formula1.Split(";")[2];
                    double currentPrice = Math.Ceiling(currentMinPrice * 1.05);
                    excelSheet.Range[$"AO{i}"].Value = currentPrice;
                    excelSheet.Range[$"AN{i}"].Value = Math.Ceiling(currentPrice * 0.125) * 10;
                }
                //если не изменилась, то удаляем эту строку
                //else
                //{
                //    ((Excel.Range)excelSheet.Rows[i]).Delete(XlDeleteShiftDirection.xlShiftUp);
                //    i--;
                //    lastRow--;
                //}
            }

            string filename = Global.DirectoryToSave 
                + $"Цены озон{shopNumber} " 
                + DateTime.Now.ToShortDateString()
                + PathToPriceTemplate.Substring(PathToPriceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }
    }
}
