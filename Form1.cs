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

        #region �������� 1� � ��������

        private void download1Cbtn_Click(object sender, EventArgs e)
        {
            builder.Set1cDataFile(GetFileName(choosedFile1c, "�������� �������� 1�"));
            invokeBalance_CheckedChanged(sender, e);
            invokePrice_CheckedChanged(sender, e);
        }

        private void downloadStrapBtn_Click(object sender, EventArgs e)
        {
            builder.SetStrapDataFile(GetFileName(choosedStrapFile, "�������� ���� � ���������"));
            invokeBalance_CheckedChanged(sender, e);
            invokePrice_CheckedChanged(sender, e);
        }

        #endregion

        #region ����
        // ��������� ������� ������ �������
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
                choosedOzon2List.Text = "�������� ������ ��������";
                builder.SetOzonList(Marketplace.Ozon3, string.Empty);
                choosedOzon3List.Text = "�������� ������ ��������";
            }
        }

        // �������� �������
        private void downloadOzon1List_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon1List, "�������� ������ ������� ��� ����1", "csv files (*.csv)|*.csv");
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
            builder.SetOzonList(Marketplace.Ozon2, GetFileName(choosedOzon2List, "�������� ������ ������� ��� ����2", "csv files (*.csv)|*.csv"));
        }

        private void downloadOzon3List_Click(object sender, EventArgs e)
        {
            builder.SetOzonList(Marketplace.Ozon3, GetFileName(choosedOzon3List, "�������� ������ ������� ��� ����3", "csv files (*.csv)|*.csv"));
        }

        // ����� ��� ������ �������� ��������
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

        // �������� �������� ��� ��������
        private void downloadOzon1Balance_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon1Balance, "�������� ������ �������� ����1");
            if (!string.IsNullOrEmpty(path))
            {
                builder.SetBalanceTemplate(Marketplace.Ozon1, path);
                addOzon1BalanceTemplate.Enabled = false;
            }
        }

        private void downloadOzon2TemplateBalance_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon2Balance, "�������� ������ �������� ����2");
            if (!string.IsNullOrEmpty(path))
            {
                builder.SetBalanceTemplate(Marketplace.Ozon2, path);
                addOzon2BalanceTemplate.Enabled = false;
            }
        }

        private void downloadOzon3TemplateBalance_Click(object sender, EventArgs e)
        {
            string path = GetFileName(choosedOzon3Balance, "�������� ������ �������� ����3");
            if (!string.IsNullOrEmpty(path))
            {
                builder.SetBalanceTemplate(Marketplace.Ozon3, path);
                addOzon3BalanceTemplate.Enabled = false;
            }
        }

        // ��������� ���������
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

        // ���������� �������� ���
        private void downloadOzon1Price_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Ozon1, GetFileName(choosedOzon1Price, "�������� ������ ��� ����1"));
            percentOzon1_ValueChanged(percentOzon1, e);
        }

        private void downloadOzon2Price_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Ozon2, GetFileName(choosedOzon2Price, "�������� ������ ��� ����2"));
            percentOzon2_ValueChanged(percentOzon2, e);
        }

        private void downloadOzon3Price_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Ozon3, GetFileName(choosedOzon3Price, "�������� ������ ��� ����3"));
            percentOzon3_ValueChanged(percentOzon3, e);
        }


        #endregion

        #region ������
        // �������
        private void downloadSelSapTemplateBalance_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Selsap, GetFileName(choosedSelSapBalance, "�������� ������ ������"));
        }

        // ����
        private void downloadSelSapTemplatePrice_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Selsap, GetFileName(choosedSelSapPrice, "�������� ������ ��� ������"));
        }
        #endregion

        #region ����������
        // ������� � ����
        private void downloadMegamarketTemplate_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Megamarket, GetFileName(choosedMegamarketFile, "�������� ������ �����������"));
            percentMegamarket_ValueChanged(percentMegamarket, e);
        }

        // �������
        private void percentMegamarket_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Megamarket, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }
        #endregion

        #region �����������
        // ������� � ����
        private void downloadAliexpressTemplate_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Aliexpress, GetFileName(choosedAliexpressFile, "�������� ������ �����������"));
            percentAliexpress_ValueChanged(percentAliexpress, e);
        }

        // �������
        private void percentAliexpress_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Aliexpress, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        #endregion

        #region ������
        // ���� ��� ������ ������� ��������
        private void addYandexBalanceTemplate_CheckedChanged(object sender, EventArgs e)
        {
            downloadYandexBalance.Enabled = choosedYandexBalance.Enabled = (sender as CheckBox)!.Checked;
        }

        // �������� ������� ��������
        private void downloadYandexBalance_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Yandex, GetFileName(choosedYandexBalance, "�������� ������ ������"));
        }

        // ��������� ��������
        private void percentYandex_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Yandex, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        // �������� ������� ���
        private void downloadYandexPrice_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Yandex, GetFileName(choosedYandexPrice, "�������� ������ ��� ������"));
            percentYandex_ValueChanged(percentYandex, e);
        }

        #endregion 

        #region ��
        // �������
        private void downloadWildberriesBalance_Click(object sender, EventArgs e)
        {
            builder.SetBalanceTemplate(Marketplace.Wildberries, GetFileName(choosedWildberriesBalance, "�������� ������ �������� ����������"));
        }

        // ����
        private void downloadWilberriesPrice_Click(object sender, EventArgs e)
        {
            builder.SetPriceTemplate(Marketplace.Wildberries, GetFileName(choosedWilberriesPrice, "�������� ������ ��� ����������"));
            percentWildberries_ValueChanged(percentWildberries, e);
        }

        // �������
        private void percentWildberries_ValueChanged(object sender, EventArgs e)
        {
            builder.SetPercent(Marketplace.Wildberries, Convert.ToDouble((sender as NumericUpDown)!.Value));
        }

        // ������ �������
        private void downloadWildberriesList_Click(object sender, EventArgs e)
        {
            builder.SetWildberriesList(GetFileName(choosedWildberriesList, "�������� ������ ������� ��� ����������"));
        }
        #endregion

        #region ��������� ��������� �����
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

        #region �������� �����/����/������
        // �����
        private void startBtn_Click(object sender, EventArgs e)
        {
            statelbl.Visible = true;
            currentState.Visible = true;
            ShowCurrentState("��������� ��������");
            startBtn.Enabled = false;
            exitBtn.Enabled = false;

            string generalPath = Global.DirectoryToSave;
            foreach (MarketplaceHandler marketplace in builder.GetMarketplaceHandlers())
            {
                if (invokeBalance.Checked)
                {
                    // ��������� ����� ��� �����������
                    Global.DirectoryToSave = Directory
                        .CreateDirectory(generalPath + $@"����� �������� {DateTime.Now.ToShortDateString()}")
                        .FullName + "\\";
                    ShowCurrentState($"���������� ������� �������� {marketplace.GetTitle}");
                    marketplace.FillBalance();
                }

                if (invokePrice.Checked)
                {
                    // ��������� ����� ��� �����������
                    Global.DirectoryToSave = Directory
                        .CreateDirectory(generalPath + $@"����� ��� {DateTime.Now.ToShortDateString()}")
                        .FullName + "\\";
                    ShowCurrentState($"���������� ������� ��� {marketplace.GetTitle}");
                    marketplace.FillPrice();
                }
            }
            Global.DirectoryToSave = generalPath;

            startBtn.Enabled = true;
            exitBtn.Enabled = true;
            currentState.Text = "���������";
        }

        //������
        private void ShowCurrentState(string message) => currentState.Text = message;

        // ����������
        private void exitBtn_Click(object sender, EventArgs e) => Close();
        #endregion

        #region ��������� ���������� ��� ���������� ������
        // ���� ��� ��������� ����������
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

        // ��������� ����������
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

        #region ��������������� �������
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
