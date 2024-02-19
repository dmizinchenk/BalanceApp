using Excel = Microsoft.Office.Interop.Excel;

namespace BalanceApp.WorkClasses.Handlers
{
    public class YandexHandlerImplement : MarketplaceHandler
    {
        public delegate void ChangePercent(Marketplace marketplace, double percent);
        public event ChangePercent? OnChangePercent;
        public new double Percent
        {
            set
            {
                base.Percent = value;
                OnChangePercent?.Invoke(Marketplace.Yandex, value);
            }
            get => base.Percent;
        }
        public override string GetTitle => "Yandex";


        public override void FillBalance()
        {
            if (PathToBalanceTemplate == null)
            {
                return;
            }

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToBalanceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Список товаров"];
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "C"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 3; i <= lastRow; i++)
            {
                int currentBalance = base.GetDataFrom1c(excelSheet, $"D{i}", $"C{i}", true);
                excelSheet.Range[$"E{i}"].Value = currentBalance;
            }

            string filename = Global.DirectoryToSave 
                + "Остатки яндекс " 
                + DateTime.Now.ToShortDateString()
                + PathToBalanceTemplate.Substring(PathToBalanceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }

        public override void FillPrice()
        {
            if(PathToPriceTemplate == null)
            {
                return;
            }

            Excel.Application excelApp = ExcelApplication.GetApp;
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(PathToPriceTemplate);
            Excel.Worksheet excelSheet = excelWorkbook.Worksheets["Список товаров"];
            int lastRow = excelSheet.Cells[excelSheet.Rows.Count, "C"].End[Excel.XlDirection.xlUp].Row;

            for (int i = 5; i <= lastRow; i++)
            {
                double priceForMarket = Math.Ceiling(base.GetDataFrom1c(excelSheet, $"D{i}", $"C{i}", false) * Percent);
                //сравниваем цену с ценой шаблона
                //если она изменилась, то записываем новые данные
                var currentPrice = excelSheet.Range[$"Q{i}"].Value;
                if (
                       //если цена есть в 1с и
                       priceForMarket > 0 &&
                       //если текущая ячейки с минимальной ценой пуста
                       (currentPrice is null ||
                       //или если она не пуста и ее значение отличается от данных 1с
                       (currentPrice is not null && Convert.ToDouble(currentPrice) != priceForMarket))
                   )
                {
                    excelSheet.Range[$"Q{i}"].Value = priceForMarket;
                    excelSheet.Range[$"R{i}"].Value = Math.Ceiling(currentPrice * 0.125) * 10;
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
                + $"Цены яндекс " 
                + DateTime.Now.ToShortDateString()
                + PathToPriceTemplate.Substring(PathToPriceTemplate.LastIndexOf('.'));
            CheckExists(filename);
            excelApp.Application.ActiveWorkbook.SaveAs(filename);

            excelWorkbook.Close();
        }
    }
}
