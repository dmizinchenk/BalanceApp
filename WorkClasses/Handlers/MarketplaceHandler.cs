using BalanceApp.WorkClasses.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.Devices;
using System.Text.RegularExpressions;

namespace BalanceApp.WorkClasses.Handlers
{
    public abstract class MarketplaceHandler
    {
        public static List<ClockModel> Clocks { set; get; } = new List<ClockModel>();
        public static List<StrapModel> Straps { set; get; } = new List<StrapModel>();
        public string PathToBalanceTemplate { get; set; } = null!;
        public string PathToPriceTemplate { get; set; } = null!;
        public double Percent { get; set; }
        public abstract string GetTitle { get; }
        public abstract void FillBalance();
        public abstract void FillPrice();
        
        protected int GetDataFrom1c(Microsoft.Office.Interop.Excel.Worksheet excelSheet, string cellWithModel, string cellWithArticle, bool isBalance)
        {
            string cellValue = Convert.ToString(excelSheet.Range[cellWithModel].Value);
            if (String.IsNullOrEmpty(cellValue))
            {
                return 0;
            }
            //если находим модель часов, то берем у них остаток
            if (cellValue.Contains("часы", StringComparison.CurrentCultureIgnoreCase))
            {
                Match? searchResult = Regex.Matches(cellValue, Global.PATTERN, RegexOptions.IgnoreCase).SingleOrDefault();
                if (searchResult is not null)
                {
                    string model = searchResult.Value.Replace(" ", String.Empty);
                    ClockModel toFind = new ClockModel() { Model = model };
                    int index = Clocks.BinarySearch(toFind);
                    if (index >= 0)
                    {
                        return isBalance ? Clocks[index].Count : Clocks[index].Price;
                    }
                    return 0;
                }
            }
            //если находим модель ремешка, то берем у него остаток
            else if (cellValue.Contains("рем", StringComparison.CurrentCultureIgnoreCase) ||
                     cellValue.Contains("бра", StringComparison.CurrentCultureIgnoreCase))
            {
                int article;
                if (excelSheet.Range[cellWithArticle].Value is not null)
                {
                    cellValue = Convert.ToString(excelSheet.Range[cellWithArticle].Value);
                    if (Int32.TryParse(cellValue.TrimStart('\''), out article))
                    {
                        StrapModel toFind = new StrapModel() { Id = article }; 
                        int index = Straps.BinarySearch(toFind);
                        if (index >= 0)
                        {
                            return isBalance ? Straps[index].Count : Straps[index].Price;
                        }
                        return 0;
                    }
                }
            }
            return 0;
        }
        protected void CheckExists(string path)
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}
