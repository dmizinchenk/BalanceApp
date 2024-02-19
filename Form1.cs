using BalanceApp.WorkClasses;
using BalanceApp.WorkClasses.Handlers;


namespace BalanceApp
{
    public partial class Form1 : Form
    {
        MarketplacesHandlerBuilder builder = new MarketplacesHandlerBuilder();
        public Form1()
        {
            InitializeComponent();
            currentDirectoryLabel.Text = Global.DirectoryToSave;
            builder.Notificator += ShowCurrentState;
            this.ControlBox = false;
        }

        #region Загрузка 1с и ремешков

        private void download1Cbtn_Click(object sender, EventArgs e)
        {
            builder.Set1cDataFile(GetFileName(choosedFile1c, "Выберите выгрузку 1С"));
            invokeBalance_CheckedChanged(sender, e);
            invokePrice_CheckedChanged(sender, e);
        }

        private void downloadStrapBtn_Click(object sender, EventArgs e)
        {
            builder.SetStrapDataFile(GetFileName(choosedStrapFile, "Выберите файл с ремешками"));
            invokeBalance_CheckedChanged(sender, e);
            invokePrice_CheckedChanged(sender, e);
        }

        #endregion

        #region Озон
        // Установка единого списка товаров
        private void ozonsSingleList_CheckedChanged(object sender, EventArgs e)
        {
            choosedOzon2List.Enabled = downloadOzon2List.Enabled = choosedOzon3List.Enabled = downloadOzon3List.Enabled = !ozonsSingleList.Checked;
            if (ozonsSingleList.Checked)
            {
                string path = (builder.GetMarketplace(Marketplace.Ozon1) as OzonHandlerImplement)!.PathToListProducts;
                builder.SetOzonList(Marketplace.Ozon2, path);
                builder.SetOzonList(Marketplace.Ozon3, path);
                choosedOzon2List.Text = choosedOzon3List.Text = choosedOzon1List.Text;
            }
            else
            {
                builder.SetOzonList(Marketplace.Ozon2, string.Empty);
                choosedOzon2List.Text = "Выберите шаблон остатков";
                builder.SetOzonList(Marketplace.Ozon3, string.Empty);
                choosedOzon3List.Text = "Выберите шаблон остатков";
            }
        }

        // Загрузка списков
        private void downloadOzon1List_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon1List, "Выберите список товаров для озон1", "csv files (*.csv)|*.csv");
            builder.SetOzonList(Marketplace.Ozon1, path);
            if (ozonsSingleList.Checked)
            {
                builder.SetOzonList(Marketplace.Ozon2, path);
                builder.SetOzonList(Marketplace.Ozon3, path);
                string[] temp = path.Split('\\');
                choosedOzon2List.Text = choosedOzon3List.Text = temp[temp.Length - 1];
            }
        }

        private void downloadOzon2List_Click(object sender, EventArgs e)
        {
            builder.SetOzonList(Marketplace.Ozon2, GetFileName(choosedOzon2List, "Выберите список товаров для озон2", "csv files (*.csv)|*.csv"));
        }

        private void downloadOzon3List_Click(object sender, EventArgs e)
        {
            builder.SetOzonList(Marketplace.Ozon3, GetFileName(choosedOzon3List, "Выберите список товаров для озон3", "csv files (*.csv)|*.csv"));
        }

        // Флаги для замены шаблонов остатков
        private void addOzon1BalanceTemplate_CheckedChanged(object sender, EventArgs e)
        {
            downloadOzon1Balance.Enabled = choosedOzon1Balance.Enabled = (sender as CheckBox)!.Checked;
        }

        private void addOzon2BalanceTemplate_CheckedChanged(object sender, EventArgs e)
        {
            downloadOzon2Balance.Enabled = choosedOzon2Balance.Enabled = (sender as CheckBox)!.Checked;
        }

        private void addOzon3BalanceTemplate_CheckedChanged(object sender, EventArgs e)
        {
            downloadOzon3Balance.Enabled = choosedOzon3Balance.Enabled = (sender as CheckBox)!.Checked;
        }

        // загрузка шаблонов для остатков
        private void downloadOzon1Balance_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon1Balance, "Выберите шаблон остатков озон1");
            if (!string.IsNullOrEmpty(path))
            {
                builder.SetBalanceTemplate(Marketplace.Ozon1, path);
                addOzon1BalanceTemplate.Enabled = false;
            }
        }

        private void downloadOzon2TemplateBalance_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon2Balance, "Выберите шаблон остатков озон2");
            if (!string.IsNullOrEmpty(path))
            {
                builder.SetBalanceTemplate(Marketplace.Ozon2, path);
                addOzon2BalanceTemplate.Enabled = false;
            }
        }

        private void downloadOzon3TemplateBalance_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon3Balance, "Выберите шаблон остатков озон3");
            if (!string.IsNullOrEmpty(path))
            {
                builder.SetBalanceTemplate(Marketplace.Ozon3, path);
                addOzon3BalanceTemplate.Enabled = false;
            }
        }

        // установка процентов
        private void percentOzon1_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Ozon1, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        private void percentOzon2_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Ozon2, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        private void percentOzon3_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Ozon3, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        // Добавление шаблонов цен
        private void downloadOzon1Price_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Ozon1, GetFileName(choosedOzon1Price, "Выберите шаблон цен озон1"));
            percentOzon1_ValueChanged(percentOzon1, e);
        }

        private void downloadOzon2Price_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Ozon2, GetFileName(choosedOzon2Price, "Выберите шаблон цен озон2"));
            percentOzon2_ValueChanged(percentOzon2, e);
        }

        private void downloadOzon3Price_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Ozon3, GetFileName(choosedOzon3Price, "Выберите шаблон цен озон3"));
            percentOzon3_ValueChanged(percentOzon3, e);
        }


        #endregion

        #region Селсап
        // Остатки
        private void downloadSelSapTemplateBalance_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Selsap, GetFileName(choosedSelSapBalance, "Выберите шаблон Селсап"));
        }

        // Цены
        private void downloadSelSapTemplatePrice_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Selsap, GetFileName(choosedSelSapPrice, "Выберите шаблон цен селсап"));
        }
        #endregion

        #region Мегамаркет
        // Остатки и цены
        private void downloadMegamarketTemplate_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Megamarket, GetFileName(choosedMegamarketFile, "Выберите шаблон мегамаркета"));
            percentMegamarket_ValueChanged(percentMegamarket, e);
        }

        // Процент
        private void percentMegamarket_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Megamarket, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }
        #endregion

        #region Алиэкспресс
        // Остатки и цены
        private void downloadAliexpressTemplate_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Aliexpress, GetFileName(choosedAliexpressFile, "Выберите шаблон Алиэкспресс"));
            percentAliexpress_ValueChanged(percentAliexpress, e);
        }

        // Процент
        private void percentAliexpress_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Aliexpress, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        #endregion

        #region Яндекс
        // Флаг для замены шаблона остатков
        private void addYandexBalanceTemplate_CheckedChanged(object sender, EventArgs e)
        {
            downloadYandexBalance.Enabled = choosedYandexBalance.Enabled = (sender as CheckBox)!.Checked;
        }

        // Загрузка шаблона остатков
        private void downloadYandexBalance_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Yandex, GetFileName(choosedYandexBalance, "Выберите шаблон Яндекс"));
        }

        // Установка процента
        private void percentYandex_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Yandex, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        // Загрузка шаблона цен
        private void downloadYandexPrice_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Yandex, GetFileName(choosedYandexPrice, "Выберите шаблон цен яндекс"));
            percentYandex_ValueChanged(percentYandex, e);
        }

        #endregion 

        #region ВБ
        // Остатки
        private void downloadWildberriesBalance_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Wildberries, GetFileName(choosedWildberriesBalance, "Выберите шаблон остатков вайлдберис"));
        }

        // Цены
        private void downloadWilberriesPrice_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Wildberries, GetFileName(choosedWilberriesPrice, "Выберите шаблон цен вайлдберис"));
            percentWildberries_ValueChanged(percentWildberries, e);
        }

        // Процент
        private void percentWildberries_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Wildberries, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        // Список товаров
        private void downloadWildberriesList_Click(object sender, EventArgs e)
        {
            builder.SetWildberriesList(GetFileName(choosedWildberriesList, "Выберите список товаров для вайлдберис"));
        }
        #endregion

        #region Установка видимости групп
        private void invokeBalance_CheckedChanged(object sender, EventArgs e)
        {
            if (builder.IsLoadDataFiles())
            {
                OzonGroupBalance.Enabled =
                    YandexGroupBalance.Enabled =
                    SelsapGroupBalance.Enabled = invokeBalance.Checked;

                MegamarketGroupBalancePrice.Enabled =
                    AliexpressGroupBalancePrice.Enabled =
                    WildberriesGroupBalance.Enabled =
                    invokeBalance.Checked || invokePrice.Checked;

            }
        }

        private void invokePrice_CheckedChanged(object sender, EventArgs e)
        {
            if (builder.IsLoadDataFiles())
            {
                OzonGroupPrice.Enabled =
                    YandexGroupPrice.Enabled =
                    SelsapGroupPrice.Enabled =
                    WildberriesGroupPrice.Enabled = invokePrice.Checked;

                MegamarketGroupBalancePrice.Enabled =
                    AliexpressGroupBalancePrice.Enabled =
                    WildberriesGroupBalance.Enabled = 
                    invokeBalance.Checked || invokePrice.Checked;
            }
        }

        private void addWildberriesTemplateBalance_CheckedChanged(object sender, EventArgs e)
        {
            downloadWildberriesBalance.Enabled = choosedWildberriesBalance.Enabled = (sender as CheckBox)!.Checked;
        }
        #endregion

        #region Процессы старт/стоп/статус
        // Старт
        private void startBtn_Click(object sender, EventArgs e)
        {
            statelbl.Visible = true;
            currentState.Visible = true;
            ShowCurrentState("Программа запущена");
            startBtn.Enabled = false;
            exitBtn.Enabled = false;

            string generalPath = Global.DirectoryToSave;
            foreach (MarketplaceHandler marketplace in builder.GetMarketplaceHandlers())
            {
                if (invokeBalance.Checked)
                {
                    // отдельная папка для результатов
                    Global.DirectoryToSave = Directory
                        .CreateDirectory(generalPath + $@"Файлы остатков {DateTime.Now.ToShortDateString()}")
                        .FullName + "\\";
                    ShowCurrentState($"Заполнение шаблона остатков {marketplace.GetTitle}");
                    marketplace.FillBalance();
                }

                if (invokePrice.Checked)
                {
                    // отдельная папка для результатов
                    Global.DirectoryToSave = Directory
                        .CreateDirectory(generalPath + $@"Файлы цен {DateTime.Now.ToShortDateString()}")
                        .FullName + "\\";
                    ShowCurrentState($"Заполнение шаблона цен {marketplace.GetTitle}");
                    marketplace.FillPrice();
                }
            }
            Global.DirectoryToSave = generalPath;

            startBtn.Enabled = true;
            exitBtn.Enabled = true;
            currentState.Text = "Выполнено";
        }

        //Статус
        private void ShowCurrentState(string message) => currentState.Text = message;

        // Завершение
        private void exitBtn_Click(object sender, EventArgs e) => Close();
        #endregion

        #region Изменение директории для сохранения файлов
        // Флаг для установки директории
        private void chooseDicrectory_CheckedChanged(object sender, EventArgs e)
        {
            curDir.Visible = !curDir.Visible;
            currentDirectoryLabel.Enabled = !currentDirectoryLabel.Enabled;
            if (!curDir.Visible)
            {
                Global.DirectoryToSave = Directory.GetCurrentDirectory() + '\\';
                currentDirectoryLabel.Text = Global.DirectoryToSave;
            }
        }

        // Установка директории
        private void curDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                Global.DirectoryToSave = fbd.SelectedPath + "\\";
                currentDirectoryLabel.Text = Global.DirectoryToSave;
            }
        }
        #endregion

        #region Вспомогательные функции
        private string GetFileName(Label label, string title, string filter = "Excel files (*.xls,*xlsx)|*.xls;*.xlsx")
        {
            string filename = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Filter = filter;
                openFileDialog.Title = title;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string[] temp = openFileDialog.FileName.Split("\\");
                    label.Text = temp[temp.Length - 1];
                    filename = openFileDialog.FileName;
                }
            }
            return filename;
        }
        #endregion
    }
}
