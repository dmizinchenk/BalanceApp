using BalanceApp.WorkClasses.Models;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class Handler1C
    {
        public string PathToStrapsFile { set; get; } = null!;
        public string PathToClocksFile { set; get; } = null!;

        public void Execute()
        {
            if (PathToStrapsFile is null || PathToClocksFile is null)
            {
                throw new Exception("Не заданы входные данные");
            }

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToStrapsFile);
            Excel.Worksheet excelSheet = excelWorkbook.ActiveSheet;
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 2; i <= lastRow; i++)
            {
                StrapModel strap = new()
                {
                    Id = Convert.ToInt32(excelSheet.Range[$"A{i}"].Value),
                    Name = (string)excelSheet.Range[$"C{i}"].Value,
                    Count = Convert.ToInt32(excelSheet.Range[$"E{i}"].Value)
                };

                if (excelSheet.Range[$"B{i}"].Value != null)
                {
                    strap.Number = Convert.ToInt32(excelSheet.Range[$"B{i}"].Value);
                }
                if (excelSheet.Range[$"F{i}"].Value != null)
                {
                    strap.Name1С = Convert.ToString(excelSheet.Range[$"F{i}"].Value);
                }

                MarketplaceHandler.Straps.Add(strap);
            }
            excelWorkbook.Close(false, false, false);
            excelApp.Workbooks.Close();

            MarketplaceHandler.Straps.Sort();

            excelWorkbook = excelApp.Workbooks.Open(PathToClocksFile);
            excelSheet = excelWorkbook.ActiveSheet;
            //вычисляем последнюю заполненную ячейку
            lastRow = excelSheet.Cells[excelSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;

            //Проходим по всем заполненным ячейкам
            for (int i = 3; i <= lastRow; i++)
            {
                string cellValue = Convert.ToString(excelSheet.Range[$"A{i}"].Value);
                if (!String.IsNullOrEmpty(cellValue))
                {
                    if (cellValue.Contains("Часы", StringComparison.CurrentCultureIgnoreCase) && !cellValue.Contains("уценка", StringComparison.CurrentCultureIgnoreCase))
                    {
                        ClockModel clock = new();

                        Match? model = Regex.Matches(cellValue, Global.PATTERN, RegexOptions.IgnoreCase).SingleOrDefault();
                        if (model != null)
                        {
                            clock.Model = model.Value.Replace(" ", String.Empty);

                            if (excelSheet.Range[$"B{i}"].Value is not null)
                                clock.Price = Convert.ToInt32(excelSheet.Range[$"B{i}"].Value);
                            else
                                clock.Price = 0;
                            if (excelSheet.Range[$"C{i}"].Value is not null)
                                clock.Count = Convert.ToInt32(excelSheet.Range[$"C{i}"].Value);
                            else
                                clock.Count = 0;

                            MarketplaceHandler.Clocks.Add(clock);
                        }
                    }
                    else
                    {
                        StrapModel? strap = MarketplaceHandler.Straps.Where(s => {
                            if (s.Name1С is not null)
                                return s.Name1С.Equals(cellValue);
                            return false;
                        }).SingleOrDefault();
                        if (strap != null)
                            strap.Price = Convert.ToInt32(excelSheet.Range[$"B{i}"].Value);
                    }
                }
            }
            excelApp.Workbooks.Close();
            MarketplaceHandler.Clocks.Sort();
        }
    }
}
