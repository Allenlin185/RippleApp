using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using ProgramMethod;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using IFilterTextReader;
using System.Drawing;

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
            public decimal CheckValue { get; set; }
            public int Files { get; set; }
        }
        private List<ProductSet> ProductList = new List<ProductSet>();
        private ProductSet Defaults;
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
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }
                string[] InfoArr = line.Split(',');
                if (InfoArr.Length < 5)
                {
                    ErrogLog.Text = "設定檔案格式不符";
                    break;
                }
                ProductSet SetRow = new ProductSet();
                SetRow.product = InfoArr[0].Trim();
                SetRow.BwLeft = Convert.ToDecimal(InfoArr[1].Trim());
                SetRow.BwRight = Convert.ToDecimal(InfoArr[2].Trim());
                SetRow.CheckValue = Convert.ToDecimal(InfoArr[3].Trim());
                SetRow.Files = Convert.ToInt32(InfoArr[4].Trim());
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

        private void RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
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
                    ProductList.Find(row => row.product == SetInfo.product).CheckValue = SetInfo.CheckValue;
                    ProductList.Find(row => row.product == SetInfo.product).Files = SetInfo.Files;
                }
                string SettingFile = @"productinfo.csv";
                File.Delete(SettingFile);
                StreamWriter sw = null;
                sw = File.AppendText(SettingFile);
                foreach(ProductSet RowInfo in ProductList)
                {
                    sw.WriteLine(
                        RowInfo.product.Trim() + "," + 
                        RowInfo.BwLeft.ToString().Trim() + "," + 
                        RowInfo.BwRight.ToString().Trim() + "," +
                        RowInfo.CheckValue.ToString().Trim() + "," +
                        RowInfo.Files.ToString().Trim()
                    );
                }
                sw.Close();
            }
            catch(Exception Ex)
            {
                ErrogLog.Text = "儲存失敗: " + Ex.Message;
            }
            finally
            {
                //LoadProduct();
            }
        }

        private void StartProcess_Click(object sender, RoutedEventArgs e)
        {
            List<ProductSet> GetFiles = new List<ProductSet>();
            string SerialNo = "";
            string Title = "";
            if (string.IsNullOrEmpty((string)SourcePath.Content))
            {
                ErrogLog.Text = "請選擇來源資料路徑";
                return;
            }
            if (string.IsNullOrEmpty((string)ResultPath.Content))
            {
                ErrogLog.Text = "請選擇判定結果資料路徑";
                return;
            }
            if (!Directory.Exists((string)SourcePath.Content))
            {
                Directory.CreateDirectory((string)SourcePath.Content);
            }
            string[] PDFFiles = Directory.GetFiles((string)SourcePath.Content, "*.pdf");
            if (PDFFiles.Length == 0)
            {
                ErrogLog.Text = "查無波紋檔";
                return;
            }
            Array.Sort(PDFFiles);
            foreach (string PDFFile in PDFFiles)
            {
                FilterReader ReadPDF = new FilterReader(PDFFile);
                string FileContent = ReadPDF.ReadToEnd();
                ReadPDF.Close();
                string[] ContentLines = Regex.Split(FileContent, Environment.NewLine);
                int ConutRow = 0;
                ProductSet FileValue = new ProductSet();
                foreach (string LineStr in ContentLines)
                {
                    ConutRow++;
                    switch (ConutRow)
                    {
                        case 4:
                            if (LineStr.Contains("βw"))
                            {
                                string[] StrArr1 = LineStr.Split('=');
                                FileValue.BwLeft = Convert.ToDecimal(StrArr1[1].Replace(@"'", "").Trim());
                            }
                            else
                            {
                                FileValue.BwLeft = 0;
                                ConutRow++;
                            }
                            break;
                        case 8:
                            if (LineStr.Contains("βw"))
                            {
                                string[] StrArr2 = LineStr.Split('=');
                                FileValue.BwRight = Convert.ToDecimal(StrArr2[1].Replace(@"'", "").Trim());
                            }
                            else
                            {
                                FileValue.BwRight = 0;
                                ConutRow++;
                            }
                            break;
                        case 28:
                            string[] StrArr3 = LineStr.Split(':');
                            SerialNo = StrArr3[1].Trim();
                            break;
                        case 29:
                            string[] StrArr4 = LineStr.Split(':');
                            if (string.IsNullOrEmpty(FileValue.product))
                            {
                                FileValue.product = StrArr4[1].Trim();
                            }
                            else
                            {
                                if (GetFiles.Exists(row => row.product != StrArr4[1].Trim()))
                                {
                                    ErrogLog.Text = PDFFile + "產品編號不符:" + StrArr4[1].Trim();
                                    return;
                                }
                            }
                            Defaults = ProductList.Find(row => row.product == FileValue.product);
                            if (Defaults == null)
                            {
                                ErrogLog.Text = PDFFile + " 產品編號不存在設定檔";
                                return;
                            }
                            if (Defaults.Files != PDFFiles.Length)
                            {
                                ErrogLog.Text = "波紋檔不足，無法分析";
                                return;
                            }
                            break;
                    }
                }
                Title = FileValue.product + Environment.NewLine + SerialNo;
                GetFiles.Add(FileValue);
                if (!ModeFile(PDFFile, FileValue, SerialNo))
                {
                    return;
                }
            }
            Excel._Workbook oWB = CreateExcelFile(Defaults, SerialNo);
            Excel._Worksheet oSheet;
            oSheet = oWB.ActiveSheet;
            SetHeader(oSheet, Title);
            int CountRow = 3;
            foreach (ProductSet FileValue in GetFiles)
            {
                SetDetail(oSheet, FileValue, Defaults, CountRow);
                CountRow++;
            }
            AvgDifference(oSheet, Defaults, CountRow);
            oWB.Save();
        }
        private Excel._Workbook CreateExcelFile(ProductSet Defaults, string SerialNo)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            
            string ResultFile = (string)ResultPath.Content + @"\" + Defaults.product + @"\" + SerialNo + ".xlsx";
            oXL = new Excel.Application();
            oXL.Visible = true;
            oWB = oXL.Workbooks.Add(Missing.Value);
            oWB.SaveAs(Filename: ResultFile);
            
            return oWB;
        }
        private void SetHeader(Excel._Worksheet oSheet, string Title)
        {
            Excel.Range oRng;
            oSheet.Cells[1, 1] = Title;
            oRng = oSheet.Range["A1", "G1"];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.Font.Color = ColorTranslator.ToOle(Color.Blue);
            oRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng = oSheet.Cells[2, 1];
            oRng.Value = "Profile Ripple";
            SetCellWidth(oRng);
            oRng = oSheet.Cells[2, 2];
            oRng.Value = "βw Left";
            SetCellWidth(oRng);
            oRng = oSheet.Cells[2, 3];
            oRng.Value = "βw Right";
            SetCellWidth(oRng);
            oRng = oSheet.Cells[2, 4];
            oRng.Value = "|βw-βb| Left";
            SetCellWidth(oRng);
            oRng = oSheet.Cells[2, 5];
            oRng.Value = "Left Result";
            SetCellWidth(oRng);
            oRng = oSheet.Cells[2, 6];
            oRng.Value = "|βw-βb| Right";
            SetCellWidth(oRng);
            oRng = oSheet.Cells[2, 7];
            oRng.Value = "Right Result";
            SetCellWidth(oRng);
        }
        private void SetDetail(Excel._Worksheet oSheet, ProductSet FileValie, ProductSet Defaults, int CountRow)
        {
            oSheet.Cells[CountRow, 1] = "OP" + (CountRow - 2);
            oSheet.Cells[CountRow, 2] = FileValie.BwLeft;
            oSheet.Cells[CountRow, 3] = FileValie.BwRight;
            oSheet.Cells[CountRow, 4] = "=ABS(B" + CountRow + "-(" + Defaults.BwLeft + "))";
            oSheet.Cells[CountRow, 6] = "=ABS(C" + CountRow + "-(" + Defaults.BwLeft + "))";
        }
        private void SetCellWidth(Excel.Range oRng)
        {
            oRng.Font.Bold = true;
            oRng.Font.Color = ColorTranslator.ToOle(Color.DarkGreen);
            oRng.Columns.AutoFit();
            oRng.ColumnWidth = oRng.ColumnWidth + 6;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        private void AvgDifference(Excel._Worksheet oSheet, ProductSet Defaults, int CountRow)
        {
            oSheet.Cells[CountRow, 1] = "Avg Difference of top " + (CountRow-3) + " OP";
            Excel.Range oRng = oSheet.Range["A"+ CountRow, "C"+ CountRow];
            oRng.Merge();
            oRng.Font.Bold = true;
            oRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oSheet.Cells[CountRow, 4] = "=AVERAGE(D3:D3)";
            oSheet.Cells[CountRow, 5] = "=IF(D13>" + Defaults.CheckValue + ",\"OK\",\"NG\")";
            oSheet.Cells[CountRow, 6] = "=AVERAGE(F3:F" + (CountRow - 1) + ")";
            oSheet.Cells[CountRow, 7] = "=IF(F13>" + Defaults.CheckValue + ",\"OK\",\"NG\")";
            oRng = oSheet.Cells[CountRow, 5];
            if (oRng.Value == "NG")
            {
                oRng.Font.Bold = true;
                oRng.Font.Size = 16;
                oRng.Font.Color = ColorTranslator.ToOle(Color.Red);
            }
            oRng = oSheet.Cells[CountRow, 7];
            if (oRng.Value == "NG")
            {
                oRng.Font.Bold = true;
                oRng.Font.Size = 16;
                oRng.Font.Color = ColorTranslator.ToOle(Color.Red);
            }
        }
        private bool ModeFile(string SourceFile, ProductSet FileValie, string SerialNo)
        {
            string result = Path.GetFileName(SourceFile);
            string DestioationPath = (string)ResultPath.Content + @"\" + FileValie.product + @"\" + SerialNo;
            try
            {
                if (!Directory.Exists(DestioationPath))
                {
                    Directory.CreateDirectory(DestioationPath);
                }
                string readerFile = Path.Combine(DestioationPath, result);
                File.Copy(SourceFile, readerFile, true);
                File.Delete(SourceFile);
                return true;
            }
            catch (IOException copyError)
            {
                ErrogLog.Text = "搬移失敗: " + copyError.Message;
                return false;
            }
        }
    }
}
