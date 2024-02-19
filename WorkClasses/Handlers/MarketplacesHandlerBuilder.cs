
using System.Runtime.CompilerServices;

namespace BalanceApp.WorkClasses.Handlers
{
    public class MarketplacesHandlerBuilder
    {
        Handler1C handler1C;
        Dictionary<Marketplace, MarketplaceHandler> marketplaces;
        public delegate void Notification(string message);
        public event Notification? Notificator;

        public MarketplacesHandlerBuilder()
        {
            handler1C = new Handler1C();
            marketplaces = new Dictionary<Marketplace, MarketplaceHandler>();
        }

        public void Set1cDataFile(string path)
        {
            handler1C.PathToClocksFile = path;
        }

        public void SetStrapDataFile(string path)
        {
            handler1C.PathToStrapsFile = path;
        }

        public bool IsLoadDataFiles() => 
            !string.IsNullOrEmpty(handler1C.PathToClocksFile) && !string.IsNullOrEmpty(handler1C.PathToStrapsFile);

        public void SetBalanceTemplate(Marketplace marketplace, string path)
        {
            GetMarketplace(marketplace).PathToBalanceTemplate = path;
        }

        public void SetPriceTemplate(Marketplace marketplace, string path)
        {
            GetMarketplace(marketplace).PathToPriceTemplate = path;
        }

        public void SetOzonList(Marketplace marketplace, string path)
        {
            (GetMarketplace(marketplace) as OzonHandlerImplement)!.PathToListProducts = path;
        }

        public void SetWildberriesList(string path)
        {
            (GetMarketplace(Marketplace.Wildberries) as WildberriesHandlerImplement)!.PathToListProducts = path;
        }

        public void SetPercent(Marketplace marketplace, double percent)
        {
            GetMarketplace(marketplace).Percent = 1 + percent / 100;
        }

        public ICollection<MarketplaceHandler> GetMarketplaceHandlers()
        {
            Notificator?.Invoke("Обработка начальных данных");
            handler1C.Execute(); 
            Notificator?.Invoke("Данные обработаны.");
            return marketplaces.Values;
        }

        public MarketplaceHandler GetMarketplace(Marketplace marketplace)
        {
            if (marketplaces.ContainsKey(marketplace))
                return marketplaces[marketplace];

            MarketplaceHandler handler = null!;
            switch (marketplace)
            {
                case Marketplace.Ozon1:
                    handler = new OzonHandlerImplement(1) { Percent = 1.0 };
                    (handler as OzonHandlerImplement)!.OnChangePercent += changeSelsapPercent;
                    break;
                case Marketplace.Ozon2:
                    handler = new OzonHandlerImplement(2) { Percent = 1.1 };
                    break;
                case Marketplace.Ozon3:
                    handler = new OzonHandlerImplement(3) { Percent = 1.0 };
                    break;
                case Marketplace.Yandex:
                    handler = new YandexHandlerImplement() { Percent = 1.05 };
                    (handler as YandexHandlerImplement)!.OnChangePercent += changeSelsapPercent;
                    break;
                case Marketplace.Megamarket:
                    handler = new MegamarketHandlerImplement() { Percent = 1.05 };
                    break;
                case Marketplace.Aliexpress:
                    handler = new AliexpressHandlerImplement() { Percent = 1.0 };
                    break;
                case Marketplace.Wildberries:
                    handler = new WildberriesHandlerImplement() { Percent = 1.1 };
                    (handler as WildberriesHandlerImplement)!.OnChangePercent += changeSelsapPercent;
                    break;
                case Marketplace.Selsap:
                    handler = new SelsapHandlerImplement() 
                    { 
                        PercentOzon1 = GetMarketplace(Marketplace.Ozon1).Percent, 
                        PercentYandex = GetMarketplace(Marketplace.Yandex).Percent, 
                        PercentWildberries = GetMarketplace(Marketplace.Wildberries).Percent
                    };
                    break;
                default:
                    throw new Exception("Неизвестный тип маркетплейса");
            }
            marketplaces.Add(marketplace, handler);
            return handler;
        }

        private void changeSelsapPercent(Marketplace marketplace, double percent)
        {
            switch (marketplace)
            {
                case Marketplace.Ozon1:
                    (GetMarketplace(Marketplace.Selsap) as SelsapHandlerImplement)!.PercentOzon1 = percent;
                    break;
                case Marketplace.Yandex:
                    (GetMarketplace(Marketplace.Selsap) as SelsapHandlerImplement)!.PercentYandex = percent;
                    break;
                case Marketplace.Wildberries:
                    (GetMarketplace(Marketplace.Selsap) as SelsapHandlerImplement)!.PercentWildberries = percent;
                    break;
                case Marketplace.Selsap:
                case Marketplace.Ozon2:
                case Marketplace.Ozon3:
                case Marketplace.Megamarket:
                case Marketplace.Aliexpress:
                default:
                    break;
            }
        }
    }
}
