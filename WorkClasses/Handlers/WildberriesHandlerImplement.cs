using BalanceApp.WorkClasses.Models;
using System.Diagnostics;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class WildberriesHandlerImplement : MarketplaceHandler
    {

        public delegate void ChangePercent(Marketplace marketplace, double percent);
        public event ChangePercent? OnChangePercent;
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
        public new double Percent
        {
            set
            {
                base.Percent = value;
                OnChangePercent?.Invoke(Marketplace.Wildberries, value);
            }
            get => base.Percent;
        }
        public string PathToListProducts { get; set; } = null!;
        public override string GetTitle => "Wildberries";

        public WildberriesHandlerImplement()
        {
            pathToBalanceTemplate = Directory.GetCurrentDirectory() + $@"\Files\Шаблон остатков ВБ.xlsx";
        }

        public override void FillBalance()
        {
            if (string.IsNullOrEmpty(PathToBalanceTemplate) || string.IsNullOrEmpty(PathToListProducts))
                return;

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToBalanceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.ActiveSheet;

            Excel.Application excelAppList = new Excel.Application();
            Process excelProc = Process.GetProcessesByName("EXCEL").Last();
            Excel.Workbook excelWorkbookList = excelApp.Workbooks.Open(PathToListProducts);
            Excel.Worksheet excelSheetList = excelWorkbook.ActiveSheet;

            int lastRowList = excelSheetList.Cells[excelSheetList.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 2; i <= lastRowList; i++)
            {
                int currentBalance = base.GetDataFrom1c(excelSheet, $"L{i}", $"A{i}", true);

                string barcode = Convert.ToString(excelSheetList.Range[$"E{i}"].Value);
                //Заполняем баркод в шаблоне
                excelSheet.Range[$"A{i}"].Value = barcode.Split(";")[0];
                //заполняем количество
                excelSheet.Range[$"B{i}"].Value = currentBalance;
                
            }
            excelAppList.Application.Quit();
            excelProc.Kill();

            string filename = Global.DirectoryToSave 
                + $"Остатки ВБ" 
                + DateTime.Now.ToShortDateString() 
                + PathToBalanceTemplate.Substring(PathToBalanceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }

        public override void FillPrice()
        {
            if (string.IsNullOrEmpty(PathToListProducts) || string.IsNullOrEmpty(PathToPriceTemplate))
                return;

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToListProducts);
            Excel.Worksheet excelSheet = excelWorkbook.ActiveSheet;

            int lastRowList = excelSheet.Cells[excelSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 2; i <= lastRowList; i++)
            {
                string cellValue = Convert.ToString(excelSheet.Range[$"L{i}"].Value);
                if(!string.IsNullOrEmpty(cellValue) && cellValue.Contains("часы", StringComparison.CurrentCultureIgnoreCase))
                {
                    Match? searchResult = Regex.Matches(cellValue, Global.PATTERN, RegexOptions.IgnoreCase).SingleOrDefault();
                    if (searchResult is not null)
                    {
                        string model = searchResult.Value.Replace(" ", string.Empty);
                        ClockModel? clock = Clocks.Where(clock => clock.Model == model).SingleOrDefault();
                        if(clock is not null)
                        {
                            clock.Id = Convert.ToInt32(excelSheet.Range[$"A{i}"].Value);
                        }
                    }
                }
            }

            excelApp.Application.Quit();

            excelApp = ExcelApplication.GetApp;
            excelWorkbook = excelApp.Workbooks.Open(PathToPriceTemplate);
            excelSheet = excelWorkbook.ActiveSheet;

            lastRowList = excelSheet.Cells[excelSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
            for (int i = 2; i <= lastRowList; i++)
            {
                int currentId = Convert.ToInt32(excelSheet.Range[$"D{i}"].Value);
                ClockModel? clock = Clocks.Where(clock => clock.Id == currentId).SingleOrDefault();
                StrapModel? strap = Straps.Where(clock => clock.Id == currentId).SingleOrDefault();
                int oldPrice = Convert.ToInt32(excelSheet.Range[$"I{i}"].Value);
                if (clock is not null && clock.Price != oldPrice)
                {
                    excelSheet.Range[$"J{i}"].Value = Math.Ceiling(clock.Price * Percent);
                }
                else if (strap is not null && strap.Price != oldPrice)
                {
                    excelSheet.Range[$"J{i}"].Value = Math.Ceiling(strap.Price * Percent); 
                }
            }

            string filename = Global.DirectoryToSave 
                + $"Цены ВБ" 
                + DateTime.Now.ToShortDateString() 
                + PathToPriceTemplate.Substring(PathToPriceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }
    }
}
