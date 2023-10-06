using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using ProgramMethod;

namespace RippleApp
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private FileMethod PGMethod = new FileMethod();
        private class ProductSet
        {
            public string product { get; set; }
            public decimal BwLeft { get; set; }
            public decimal BwRight { get; set; }
        }
        private List<ProductSet> ProductList = new List<ProductSet>();
        public MainWindow()
        {
            InitializeComponent();
            SourcePath.Content = PGMethod.GetConfigSetting("SourcePath");
            ResultPath.Content = PGMethod.GetConfigSetting("ResultPath");
            LoadProduct();
        }
        private void LoadProduct()
        {
            ErrogLog.Text = "";
            string SettingFile = @"productinfo.csv";
            bool fileExist = File.Exists(SettingFile);
            if (!fileExist)
            {
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(SettingFile);
                    sw.WriteLine("");
                }
                catch (Exception Ex)
                {
                    ErrogLog.Text = "寫檔失敗: " + Ex.Message;
                    return;
                }
                finally
                {
                    sw.Close();
                }
            }
            StreamReader reader = new StreamReader(File.OpenRead(SettingFile));
            ProductList.Clear();
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                string[] InfoArr = line.Split(',');
                if (InfoArr.Length < 3)
                {
                    ErrogLog.Text = "設定檔案格式不符";
                    break;
                }
                ProductSet SetRow = new ProductSet();
                SetRow.product = InfoArr[0].Trim();
                SetRow.BwLeft = Convert.ToDecimal(InfoArr[1].Trim());
                SetRow.BwRight = Convert.ToDecimal(InfoArr[2].Trim());
                ProductList.Add(SetRow);
            }
            ItemsSetting.ItemsSource = ProductList;
        }
        private void SetSouce_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog GetFolder = new FolderBrowserDialog
            {
                Description = "請選擇來源資料夾"
            };
            if (GetFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SourcePath.Content = GetFolder.SelectedPath;
                PGMethod.SetConfigSetting("SourcePath", GetFolder.SelectedPath);
            }
        }

        private void SetResult_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog GetFolder = new FolderBrowserDialog
            {
                Description = "請選擇判定結果資料夾"
            };
            if (GetFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ResultPath.Content = GetFolder.SelectedPath;
                PGMethod.SetConfigSetting("ResultPath", GetFolder.SelectedPath);
            }
        }

        private void RowEditEnding(object sender, System.Windows.Controls.DataGridRowEditEndingEventArgs e)
        {
            ProductSet SetInfo = e.Row.Item as ProductSet;
            try
            {
                ProductSet FindRow = ProductList.Find(row => row.product == SetInfo.product);
                if (FindRow == null)
                {
                    ProductList.Add(SetInfo);
                }
                else
                {
                    ProductList.Find(row => row.product == SetInfo.product).BwLeft = SetInfo.BwLeft;
                    ProductList.Find(row => row.product == SetInfo.product).BwRight = SetInfo.BwRight;
                }
                string SettingFile = @"productinfo.csv";
                File.Delete(SettingFile);
                StreamWriter sw = null;
                sw = File.AppendText(SettingFile);
                foreach(ProductSet RowInfo in ProductList)
                {
                    sw.WriteLine(RowInfo.product.Trim() + "," + RowInfo.BwLeft.ToString().Trim() + "," + RowInfo.BwRight.ToString().Trim());
                }
                sw.Close();
            }
            catch(Exception Ex)
            {
                ErrogLog.Text = "儲存失敗: " + Ex.Message;
            }
            finally
            {
                LoadProduct();
            }
        }
    }
}
