﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Runtime.InteropServices;
using System.Data.SQLite;
using Dapper;
using System.Globalization;
using SpreadsheetLight.Charts;
using System.Collections;

namespace MeasuresData
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
        }
        #region

        public class MMeasure
        {
            public string Group { get; set; }
            public string MeasureID { get; set; }
            public string MeasureName { get; set; }
            public string Numerator { get; set; }
            public string NumeratorName { get; set; }
            public string Denominator { get; set; }
            public string DenominatorName { get; set; }
            public string Threshold { get; set; }
            public string Frequency { get; set; }
            public bool Positive { get; set; }
            public string User { get; set; }
            public Dictionary<string, List<double>> Records = new Dictionary<string, List<double>>();
        }

        public class MElement
        {
            public string Group { get; set; }
            public string Depart { get; set; }
            public string ElementID { get; set; }
            public string ElementName { get; set; }
            public string Frequency { get; set; }
            public string User { get; set; }
            public bool Complex { get; set; }

            public List<string> SameID = new List<string>();

            public DateTime ElementDate { get; set; }
            public string ElementRecord { get; set; }
            public Dictionary<string, string> PreDdatas { get; set; }
            public MElement()
            {
                PreDdatas = new Dictionary<string, string>();
                ElementDate = new DateTime();
            }
        }

        public class FileStatusHelper
        {
            [DllImport("kernel32.dll")]
            public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

            [DllImport("kernel32.dll")]
            public static extern bool CloseHandle(IntPtr hObject);

            public const int OF_READWRITE = 2;
            public const int OF_SHARE_DENY_NONE = 0x40;
            public static readonly IntPtr HFILE_ERROR = new IntPtr(-1);

            /// <summary>
            /// 查看檔案是否被佔用
            /// </summary>
            /// <param name="filePath"></param>
            /// <returns></returns>
            public static bool IsFileOccupied(string filePath)
            {
                IntPtr vHandle = _lopen(filePath, OF_READWRITE | OF_SHARE_DENY_NONE);
                CloseHandle(vHandle);
                return vHandle == HFILE_ERROR ? true : false;
            }
        }
        private Dictionary<string, double> StSummary(List<double> ListData)
        {
            Dictionary<string, double> stSummary = new Dictionary<string, double>();

            double _value = 0;
            double Mean = 0;//平均值
            double Sum = 0;//總和
            double StandardDeviation = 0;//樣本標準差
            double Variance = 0;//變異數
            double Median = 0;//中位數
            double FirstQuartile = 0;//第一四分位數
            double ThirdQuartile = 0;//第三四分位數
            double IQR = 0;//四分位距
            double Minimum = double.MaxValue;//最小值
            double Maximum = double.MinValue;//最大值
            double Range = 0;//全距
            double SumOfSquare = 0;//平方總和

            //計算值取小數點8位
            if (ListData.Count > 0)
            {
                for (int i = 0; i < ListData.Count; i++)
                {
                    _value = ListData[i];
                    if (Maximum < _value) Maximum = _value;
                    if (Minimum > _value) Minimum = _value;
                    Sum += _value;
                }

                Mean = Math.Round(Sum / ListData.Count, 8);
                Sum = Math.Round(Sum, 8);
                Range = Math.Round(Maximum - Minimum, 8);
                Maximum = Math.Round(Maximum, 8);
                Minimum = Math.Round(Minimum, 8);

                for (int i = 0; i < ListData.Count; i++)
                {
                    _value = ListData[i];
                    SumOfSquare += Math.Pow((_value - Mean), 2);
                }

                if (ListData.Count > 1)
                {
                    Variance = SumOfSquare / (ListData.Count - 1);
                    StandardDeviation = Math.Sqrt(SumOfSquare / (ListData.Count - 1));
                    Variance = Math.Round(Variance, 8);
                    StandardDeviation = Math.Round(StandardDeviation, 8);
                }
                else
                {
                    Variance = double.NaN;
                    StandardDeviation = double.NaN;
                }

                Median = Math.Round(GetMedian(ListData), 8);

                if (ListData.Count > 1)
                {
                    FirstQuartile = GetFirstQuartile(ListData);
                    ThirdQuartile = GetThirdQuartile(ListData);
                }
                else
                {
                    FirstQuartile = ListData[0];
                    ThirdQuartile = ListData[0];
                }

                IQR = ThirdQuartile - FirstQuartile;
                FirstQuartile = Math.Round(FirstQuartile, 8);
                ThirdQuartile = Math.Round(ThirdQuartile, 8);
                IQR = Math.Round(IQR, 8);
            }
            else
            {
                Mean = double.NaN;
                Sum = double.NaN;
                StandardDeviation = double.NaN;
                Variance = double.NaN;
                Median = double.NaN;
                FirstQuartile = double.NaN;
                ThirdQuartile = double.NaN;
                IQR = double.NaN;
                Minimum = double.NaN;
                Maximum = double.NaN;
                Range = double.NaN;
            }

            stSummary.Add("Mean", Mean);
            stSummary.Add("Sum", Sum);
            stSummary.Add("StandardDeviation", StandardDeviation);
            stSummary.Add("Variance", Variance);
            stSummary.Add("Median", Median);
            stSummary.Add("FirstQuartile", FirstQuartile);
            stSummary.Add("ThirdQuartile", ThirdQuartile);
            stSummary.Add("IQR", IQR);
            stSummary.Add("Minimum", Minimum);
            stSummary.Add("Maximum", Maximum);
            stSummary.Add("Range", Range);

            return stSummary;
        }
        private double GetMedian(List<double> ListData)
        {
            double _value = 0;

            if (ListData.Count % 2 == 0)//數量為偶數
            {
                int _index = ListData.Count / 2;
                double valLeft = (double)ListData[_index - 1];
                double valRight = (double)ListData[_index];
                _value = (valLeft + valRight) / 2;
            }
            else//數量為奇數
            {
                int _index = (ListData.Count + 1) / 2;
                _value = (double)ListData[_index - 1];
            }

            return _value;
        }
        private double GetFirstQuartile(List<double> ListData)
        {
            double _value = 0;
            double _firstQuartilePosition = (double)(ListData.Count + 3) / 4;
            double _lowerValue = ListData[(int)Math.Floor(_firstQuartilePosition) - 1];
            double _upperValue = ListData[(int)Math.Floor(_firstQuartilePosition)];
            double _factor = _firstQuartilePosition - Math.Floor(_firstQuartilePosition);
            _value = _lowerValue + _factor * (_upperValue - _lowerValue);

            return _value;
        }

        private double GetThirdQuartile(List<double> ListaData)
        {
            double _value = 0;
            double _firstQuartilePosition = (3 * (double)ListaData.Count + 1) / 4;
            double _lowerValue = ListaData[(int)Math.Floor(_firstQuartilePosition) - 1];
            double _upperValue = ListaData[(int)Math.Floor(_firstQuartilePosition)];
            double _factor = _firstQuartilePosition - Math.Floor(_firstQuartilePosition);
            _value = _lowerValue + _factor * (_upperValue - _lowerValue);

            return _value;
        }
        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }
        //原始資料庫之要素資料
        public List<MElement> gdata = new List<MElement>();
        //原始資料庫之指標資料
        public List<MMeasure> gmeasure = new List<MMeasure>();
        //收回單位填寫之要素數值
        public List<MElement> gcollect = new List<MElement>();
        //資料庫中過往要素數值
        public Dictionary<string, List<MElement>> gbackups = new Dictionary<string, List<MElement>>();
        //原始資料庫之重覆要素
        public Dictionary<string, List<string>> gduplicate = new Dictionary<string, List<string>>();

        public List<string> SameEle = new List<string>();

        public void LoadFile(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return;
            gdata.Clear();
            gmeasure.Clear();
            gduplicate.Clear();
            SameEle.Clear();
            gbackups.Clear();
            try
            {
                SLDocument sl = new SLDocument(fname, "工作表1");

                for (int i = 0; i < 500; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                        continue;
                    MElement data = new MElement
                    {
                        Group = sl.GetCellValueAsString(i + 2, 1).Trim(),
                        Depart = sl.GetCellValueAsString(i + 2, 2).Trim(),
                        ElementID = sl.GetCellValueAsString(i + 2, 3).Trim(),
                        ElementName = sl.GetCellValueAsString(i + 2, 5).Trim(),
                        User = sl.GetCellValueAsString(i + 2, 7).Trim()
                    };
                    for (int j = 12; j < 15; j++)
                    {
                        string content = sl.GetCellValueAsString(i + 2, j).Trim();
                        if (string.IsNullOrEmpty(content))
                            break;

                        data.Complex = content.Contains("+");

                        if (SameEle.Contains(content) || (gduplicate.ContainsKey(content) && gduplicate[content].Contains(data.ElementID)))
                            continue;

                        SameEle.Add(content);

                        if (gduplicate.ContainsKey(data.ElementID))
                        {
                            if (!gduplicate[data.ElementID].Contains(content))
                            {
                                gduplicate[data.ElementID].Add(content);
                            }
                        }
                        else
                        {
                            gduplicate.Add(data.ElementID, new List<string>() { content });
                        }
                    }

                    gdata.Add(data);
                }
                sl.CloseWithoutSaving();

                if (gdata.Count > 0)
                {
                    MessageBox.Show("匯入成功 : " + gdata.Count.ToString());
                    TxtBox1.Text += Environment.NewLine + "指標匯入數量 : " + gdata.Count + Environment.NewLine;
                    if (gduplicate.Count > 0)
                    {
                        TxtBox1.Text += Environment.NewLine + "相同意義要素組數量 : " + gduplicate.Count + Environment.NewLine;
                    }
                }
                else
                {
                    MessageBox.Show("匯入失敗");
                }

                SLDocument sl2 = new SLDocument(fname, "工作表2");
                for (int i = 0; i < 500; i++)
                {
                    if (string.IsNullOrEmpty(sl2.GetCellValueAsString(i + 2, 1)))
                        break;
                    if (string.IsNullOrEmpty(sl2.GetCellValueAsString(i + 2, 2)))
                        continue;
                    MMeasure data = new MMeasure
                    {
                        Group = sl2.GetCellValueAsString(i + 2, 1).Trim(),
                        MeasureID = sl2.GetCellValueAsString(i + 2, 2).Trim(),
                        MeasureName = sl2.GetCellValueAsString(i + 2, 3).Trim(),
                        Numerator = sl2.GetCellValueAsString(i + 2, 4).Trim(),
                        NumeratorName = sl2.GetCellValueAsString(i + 2, 5).Trim(),
                        Denominator = sl2.GetCellValueAsString(i + 2, 6).Trim(),
                        DenominatorName = sl2.GetCellValueAsString(i + 2, 7).Trim(),
                    };

                    gmeasure.Add(data);
                }

                sl2.CloseWithoutSaving();

                MessageBox.Show("匯入成功 : " + gmeasure.Count.ToString());

                if (gmeasure.Count > 0)
                {
                    List<string> group = new List<string>() { "全部" };
                    foreach (var x in gmeasure)
                    {
                        if (group.Contains(x.Group))
                            continue;
                        else
                        {
                            group.Add(x.Group);
                        }
                    }
                    Combx1.ItemsSource = group;
                    Combx2.ItemsSource = gmeasure.Select(o => o.MeasureID + ":" + o.MeasureName).ToList();
                    Combx1.SelectedIndex = 0;
                    List<string> spctype = new List<string>() { "Default", "U", "C", "P", "I", "nP", "Xbar_S", "Xbar_R" };
                    Combx3.ItemsSource = spctype;
                    Combx3.SelectedIndex = 0;
                }
                LoadDataBASE();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDataBASE()
        {
            string fpath = Environment.CurrentDirectory + @"\要素備份";
            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = fpath + @"\指標收集存檔總檔.xlsx";
            if (!System.IO.File.Exists(fname))
                return;
            using (SLDocument sl = new SLDocument(fname))
            {
                for (int i = 0; i < 500; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;

                    if (gbackups.Count > 0 && gbackups.ContainsKey(sl.GetCellValueAsString(i + 2, 1)))
                    {
                        continue;
                    }
                    List<MElement> lme = new List<MElement>();
                    List<string> duplicate = new List<string>();
                    for (int j = 0; j < 12 + 6; j++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(1, j + 2)))
                            break;
                        if (!DateTime.TryParse(sl.GetCellValueAsString(1, j + 2), out DateTime dts))
                            break;
                        if (dts > DateTime.Now.AddMonths(-1 - 12 - 6) && dts < DateTime.Now)
                        {
                            if (duplicate.Contains(dts.ToString("yyyy/MM")))
                                continue;
                            else
                                duplicate.Add(dts.ToString("yyyy/MM"));
                            if (!string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, j + 2)))
                            {
                                MElement data = new MElement
                                {
                                    ElementID = sl.GetCellValueAsString(i + 2, 1).Trim(),
                                    ElementRecord = sl.GetCellValueAsString(i + 2, j + 2).Trim(),
                                    ElementDate = dts
                                };
                                lme.Add(data);
                            }
                        }
                    }
                    gbackups[sl.GetCellValueAsString(i + 2, 1)] = lme;

                    try
                    {
                        if (gduplicate.ContainsKey(sl.GetCellValueAsString(i + 2, 1)))
                        {
                            var glists = gduplicate.Where(o => o.Key == sl.GetCellValueAsString(i + 2, 1)).FirstOrDefault().Value;
                            foreach (var x in glists)
                            {
                                if (!gbackups.ContainsKey(x) &&
                                    gdata.Find(o => o.ElementID == x) != null)
                                {
                                    gbackups[x] = lme;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }
                }

                sl.CloseWithoutSaving();
            }
        }

        private void BT_IMPORT_SOURCE(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = Environment.CurrentDirectory;
            dlg.Title = "選取資料檔";
            dlg.Filter = "xlsx files (*.*)|*.xlsx";
            if (dlg.ShowDialog() == true)
            {
                LoadFile(dlg.FileName);
            }
        }
        public void ExportData(string type)
        {
            if (!Directory.Exists(System.Environment.CurrentDirectory + @"\資料上傳"))
            {
                Directory.CreateDirectory(System.Environment.CurrentDirectory + @"\資料上傳");
            }
            if (gdata.Count <= 0)
            {
                MessageBox.Show("尚未匯入清單");
                return;
            }
            if (type == "TCPI")
            {
                SLDocument nsl = new SLDocument();
                /*
                SLStyle style = nsl.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                style.Alignment.WrapText = true;
                style.Font.FontSize = 12;
                */
                nsl.SetColumnWidth(1, 10);
                nsl.SetColumnWidth(2, 15);
                nsl.SetColumnWidth(3, 45);
                nsl.SetColumnWidth(4, 10);
                nsl.SetColumnWidth(5, 15);
                nsl.SetCellValue(1, 1, "日期");
                nsl.SetCellValue(1, 2, "要素代碼");
                nsl.SetCellValue(1, 3, "要素名稱");
                nsl.SetCellValue(1, 4, "頻率");
                nsl.SetCellValue(1, 5, "提報要素值");
                if (gdata.Count > 0)
                {
                    int index = 2;
                    for (int i = 0; i < gdata.Count; i++)
                    {
                        if (gdata[i].Group != type)
                            continue;
                        nsl.SetCellValue(index, 1, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                        nsl.SetCellValue(index, 2, gdata[i].ElementID);
                        nsl.SetCellValue(index, 3, gdata[i].ElementName);
                        nsl.SetCellValue(index, 4, "月");
                        if (gcollect.Count > 0)
                        {
                            var number = from num in gcollect
                                         where num.ElementID == gdata[i].ElementID && !string.IsNullOrEmpty(num.ElementRecord)
                                         select num;
                            if (number != null && number.ToList().Count > 0)
                                nsl.SetCellValue(index, 5, Convert.ToDouble(number.ToList().First().ElementRecord));
                        }
                        index++;
                    }
                }
                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                nsl.SaveAs(Refile);
            }
            else if (type == "評鑑持續")
            {
                /*
                if (!File.Exists(System.Environment.CurrentDirectory + @"\201905.xlsx"))
                    return;
                SLDocument nsl = new SLDocument(System.Environment.CurrentDirectory + @"\201905.xlsx");
                Random rand = new Random();
                for (int i = 2; i < 94; i++)
                {
                    nsl.SetCellValue(i, 3, "55" + rand.Next(10, 200));
                }
                nsl.Save();
                */

                SLDocument nsl = new SLDocument();
                SLStyle style = nsl.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                style.Font.Bold = true;
                style.Font.FontSize = 12;
                style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 176, 240), System.Drawing.Color.Black);
                SLStyle style2 = nsl.CreateStyle();
                style2.Alignment.WrapText = true;
                style2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style2.Alignment.Vertical = VerticalAlignmentValues.Center;

                nsl.SetColumnStyle(2, style2);
                nsl.SetCellStyle(1, 1, style);
                nsl.SetCellStyle(1, 2, style);
                nsl.SetCellStyle(1, 3, style);
                nsl.SetCellStyle(1, 4, style);
                nsl.SetCellStyle(1, 5, style);
                nsl.SetCellStyle(1, 6, style);
                nsl.SetCellStyle(1, 7, style);
                nsl.SetCellStyle(1, 8, style);
                nsl.SetColumnWidth(1, 15);
                nsl.SetColumnWidth(2, 45);
                nsl.SetColumnWidth(3, 15);
                nsl.SetColumnWidth(4, 15);
                nsl.SetColumnWidth(5, 15);
                nsl.SetColumnWidth(6, 15);
                nsl.SetColumnWidth(7, 15);
                nsl.SetColumnWidth(8, 15);
                nsl.SetCellValue(1, 1, "要素代碼");
                nsl.SetCellValue(1, 2, "要素名稱");
                nsl.SetCellValue(1, 3, DateTime.Now.AddMonths(-1).ToString("yyyy/MM") + "(月)");
                nsl.SetCellValue(1, 4, DateTime.Now.AddMonths(-2).ToString("yyyy/MM") + "(月)");
                nsl.SetCellValue(1, 5, DateTime.Now.AddMonths(-3).ToString("yyyy/MM") + "(月)");
                nsl.SetCellValue(1, 6, DateTime.Now.AddMonths(-4).ToString("yyyy/MM") + "(月)");
                nsl.SetCellValue(1, 7, DateTime.Now.AddMonths(-5).ToString("yyyy/MM") + "(月)");
                nsl.SetCellValue(1, 8, DateTime.Now.AddMonths(-6).ToString("yyyy/MM") + "(月)");

                if (gdata.Count > 0)
                {
                    var sortdata = gdata;
                    sortdata.Sort((x, y) => { return x.ElementID.CompareTo(y.ElementID); });
                    int index = 2;
                    for (int i = 0; i < sortdata.Count; i++)
                    {
                        if (sortdata[i].Group != type)
                            continue;
                        nsl.SetCellValue(index, 1, sortdata[i].ElementID);
                        nsl.SetCellValue(index, 2, sortdata[i].ElementName);
                        if (gcollect.Count > 0)
                        {
                            var number = from num in gcollect
                                         where num.ElementID == gdata[i].ElementID && !string.IsNullOrEmpty(num.ElementRecord)
                                         select num;
                            if (number != null && number.ToList().Count > 0)
                                nsl.SetCellValue(index, 3, Convert.ToDouble(number.ToList().First().ElementRecord));
                        }
                        index++;
                    }
                }

                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                nsl.SaveAs(Refile);

            }
            else if (type == "THIS")
            {
                SLDocument nsl = new SLDocument();
                /*
                SLStyle style = nsl.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                style.Alignment.WrapText = true;
                style.Font.FontSize = 12;
                */
                nsl.SetColumnWidth(1, 15);
                nsl.SetColumnWidth(2, 15);
                nsl.SetColumnWidth(3, 10);
                nsl.SetColumnWidth(4, 10);
                nsl.SetColumnWidth(5, 10);
                nsl.SetColumnWidth(6, 10);
                nsl.SetCellValue(1, 1, "醫院會員代碼");
                nsl.SetCellValue(1, 2, "提報民國年分");
                nsl.SetCellValue(1, 3, "提報月份");
                nsl.SetCellValue(1, 4, "指標代碼");
                nsl.SetCellValue(1, 5, "分子數據");
                nsl.SetCellValue(1, 6, "分母數據");

                if (gmeasure.Count > 0)
                {
                    int index = 2;
                    Random rand = new Random();
                    for (int i = 0; i < gmeasure.Count; i++)
                    {
                        if (gmeasure[i].Group != type)
                            continue;
                        nsl.SetCellValue(index, 1, "JB0005");
                        nsl.SetCellValue(index, 2, DateTime.Now.AddMonths(-1).AddYears(-1911).Year);
                        nsl.SetCellValue(index, 3, DateTime.Now.AddMonths(-1).Month);
                        nsl.SetCellValue(index, 4, gmeasure[i].MeasureID);
                        if (gcollect.Count > 0)
                        {
                            if (gmeasure[i].Numerator == "1")
                                nsl.SetCellValue(index, 5, "1");
                            else
                            {
                                var numer = gcollect.FirstOrDefault(o => o.ElementID == gmeasure[i].Numerator && !string.IsNullOrEmpty(o.ElementRecord));
                                if (numer != null)
                                    nsl.SetCellValue(index, 5, Convert.ToDouble(numer.ElementRecord));
                            }
                            if (gmeasure[i].Denominator == "1")
                                nsl.SetCellValue(index, 6, "1");
                            else
                            {
                                var deno = gcollect.FirstOrDefault(o => o.ElementID == gmeasure[i].Denominator && !string.IsNullOrEmpty(o.ElementRecord));
                                if (deno != null)
                                    nsl.SetCellValue(index, 6, Convert.ToDouble(deno.ElementRecord));
                            }
                            /*
                            var numer = from num in gcollect
                                        where num.Element == gmeasure[i].Numerator && !string.IsNullOrEmpty(num.ElementData)
                                        select num;
                            if (numer != null && numer.ToList().Count > 0)
                                nsl.SetCellValue(index, 5, numer.ToList().First().ElementData);
                            var deno = from num in gcollect
                                       where num.Element == gmeasure[i].Denominator && !string.IsNullOrEmpty(num.ElementData)
                                       select num;
                            if (deno != null && deno.ToList().Count > 0)
                                nsl.SetCellValue(index, 6, deno.ToList().First().ElementData);
                           */

                        }
                        //nsl.SetCellValue(index, 5, gmeasure[i].Numerator);
                        //nsl.SetCellValue(index, 6, gmeasure[i].Denominator);
                        index++;
                    }
                }
                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                nsl.SaveAs(Refile);
            }
            MessageBox.Show("轉檔成功");
        }

        private void BT_TO_TCPI(object sender, RoutedEventArgs e)
        {
            ExportData("TCPI");

            /*
            if (FileStatusHelper.IsFileOccupied(@"檔案路徑"))
            {
                MessageBox.Show("檔案已被佔用");
            }
            else
            {
                MessageBox.Show("檔案未被佔用");
            }
            */
        }

        private void BT_TO_HACMI(object sender, RoutedEventArgs e)
        {
            ExportData("評鑑持續");
        }

        private void BT_TO_THIS(object sender, RoutedEventArgs e)
        {
            ExportData("THIS");
        }

        private void BT_IMPORT_MEASUREDATA(object sender, RoutedEventArgs e)
        {
            if (!System.IO.File.Exists(System.Environment.CurrentDirectory + @"\measurements.db"))
                return;
            string dbPath = System.Environment.CurrentDirectory + @"\measurements.db";
            string cnStr = "Data Source=" + dbPath + ";Version=3;";

            using (var cn = new SQLiteConnection(cnStr))
            {
                cn.Open();

                var list = cn.Query<MElement>(
                    "SELECT * FROM MeasureTable WHERE MeasureID=@catg", new { catg = "EDP005-01" });

                foreach (var item in list)
                {
                    TxtBox1.Text += item.ElementDate.Year + "/" + item.ElementDate.Month + "=" + item.ElementRecord + Environment.NewLine;
                }

            }

        }
        private void BT_IMPORT_RESULT(object sender, RoutedEventArgs e)
        {
            gcollect.Clear();

            List<string> duplicate = new List<string>();

            string folderName = System.Environment.CurrentDirectory + @"\資料匯總";
            try
            {
                foreach (var finame in System.IO.Directory.GetFileSystemEntries(folderName))
                {
                    if (System.IO.Path.GetExtension(finame) != ".xlsx")
                        continue;
                    SLDocument sl = new SLDocument(finame);
                    if (!DateTime.TryParse(sl.GetCellValueAsString(1, 5), out DateTime time))
                        continue;
                    // 只匯入&處理前一個月的數據
                    if (time.Year != DateTime.Now.AddMonths(-1).Year
                        || time.Month != DateTime.Now.AddMonths(-1).Month)
                        continue;
                    for (int i = 2; i < 500; i++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 1)))
                            break;
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 3)))
                            continue;
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 5)))
                            continue;
                        // 匯入資料若無法轉換成double，表示資料有錯誤
                        if (!Double.TryParse(sl.GetCellValueAsString(i, 5).Trim(), out double rd))
                        {
                            MessageBox.Show("指標: " + sl.GetCellValueAsString(i, 3).Trim() + "之數值有誤 (" + sl.GetCellValueAsString(i, 5).Trim() + ")");
                            continue;
                        }
                        MElement data = new MElement
                        {
                            ElementID = sl.GetCellValueAsString(i, 3).Trim(),
                            ElementRecord = sl.GetCellValueAsString(i, 5).Trim(),
                            ElementDate = time
                        };
                        if (gcollect.Count > 0)
                        {
                            bool dup = false;
                            foreach (var x in gcollect)
                            {
                                if (data.ElementID == x.ElementID)
                                {
                                    duplicate.Add(x.ElementID);
                                    dup = true;
                                    break;
                                }
                            }
                            if (!dup)
                                gcollect.Add(data);
                        }
                        else
                            gcollect.Add(data);
                        /*
                         * 處理複合要素
                         *
                        try
                        {
                            foreach (var x in gdata)
                            {
                                if (gcollect.FirstOrDefault(o => o.ElementID == x.ElementID) != null)
                                    continue;
                                if (x.Complex && gduplicate.ContainsKey(x.ElementID))
                                {
                                    var complexs = gduplicate[x.ElementID].FirstOrDefault(o => o.Contains("+"));
                                    if (complexs == null)
                                        continue;
                                    var elements = complexs.Split('+').ToList();

                                    if (elements.Count > 0)
                                    {
                                        double total = 0;
                                        bool collected = true;
                                        foreach (var ele in elements)
                                        {
                                            var em = gcollect.FirstOrDefault(o => o.ElementID == ele.Trim());
                                            if (em == null)
                                            {
                                                collected = false;
                                                break;
                                            }
                                            if (em != null && double.TryParse(em.ElementRecord, out double record))
                                                total += record;
                                        }
                                        if (collected)
                                        {
                                            gcollect.Add(new MElement
                                            {
                                                ElementID = x.ElementID,
                                                ElementRecord = total.ToString(),
                                                ElementDate = time
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        */
                        /*
                        try
                        {
                            if (gduplicate.ContainsKey(data.ElementID))
                            {
                                var glists = gduplicate.Where(o => o.Key == data.ElementID).FirstOrDefault().Value;
                                foreach (var x in glists)
                                {
                                    if (gcollect.Find(o => o.ElementID == x.ToString()) == null &&
                                        gdata.Find(o => o.ElementID == x.ToString()) != null)
                                    {
                                        //MessageBox.Show(x.ToString());
                                        MElement dupdata = new MElement
                                        {
                                            ElementID = x.ToString(),
                                            ElementRecord = data.ElementRecord,
                                            ElementDate = data.ElementDate
                                        };
                                        gcollect.Add(dupdata);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);
                        }
                        */
                    }

                    sl.CloseWithoutSaving();
                }
                try
                {
                    if (gcollect.Count > 0)
                    {
                        /*
                         * 處理複合要素
                         */
                        /*
                       foreach (var x in gdata)
                       {
                           if (gcollect.FirstOrDefault(o => o.ElementID == x.ElementID) != null)
                               continue;
                           if (x.Complex && gduplicate.ContainsKey(x.ElementID))
                           {
                               var complexs = gduplicate[x.ElementID].FirstOrDefault(o => o.Contains("+"));
                               if (complexs == null)
                                   continue;
                               var elements = complexs.Split('+').ToList();

                               if (elements.Count > 0)
                               {
                                   double total = 0;
                                   bool collected = true;
                                   foreach (var ele in elements)
                                   {
                                       var em = gcollect.FirstOrDefault(o => o.ElementID == ele.Trim());
                                       if (em == null)
                                       {
                                           collected = false;
                                           break;
                                       }
                                       if (em != null && double.TryParse(em.ElementRecord, out double record))
                                           total += record;
                                   }
                                   if (collected)
                                   {
                                       gcollect.Add(new MElement
                                       {
                                           ElementID = x.ElementID,
                                           ElementRecord = total.ToString(),
                                           ElementDate = DateTime.Now.AddMonths(-1)
                                       });
                                   }
                               }
                           }
                           */
                        /// 處理重覆要素，執行三次以避免多重組合要性遺漏 (但需考慮效能)
                        /// 目前只能處理相加之要素
                        for (int i = 0; i < 3; i++)
                        {
                            foreach (var x in gduplicate)
                            {
                                var glists = gcollect.FirstOrDefault(o => o.ElementID == x.Key);
                                if (glists != null)
                                    continue;
                                var complexs = x.Value.FirstOrDefault(o => o.Contains("+"));
                                if (complexs != null)
                                {
                                    var elements = complexs.Split('+').ToList();
                                    bool collected = false;
                                    double total = 0;

                                    if (elements.Count > 0)
                                    {
                                        foreach (var ele in elements)
                                        {
                                            var em = gcollect.FirstOrDefault(o => o.ElementID == ele.Trim());
                                            if (em == null)
                                            {
                                                collected = false;
                                                break;
                                            }
                                            if (em != null && double.TryParse(em.ElementRecord, out double record))
                                                total += record;
                                            collected = true;
                                        }
                                    }
                                    if (collected)
                                    {
                                        gcollect.Add(new MElement
                                        {
                                            ElementID = x.Key,
                                            ElementRecord = total.ToString(),
                                            ElementDate = DateTime.Now.AddMonths(-1)
                                        });
                                        if (x.Value.Count > 1)
                                        {
                                            x.Value.ForEach((o) =>
                                            {
                                                if (!o.Contains("+") && gcollect.FirstOrDefault(z => z.ElementID == o) == null)
                                                {
                                                    //MessageBox.Show(o);
                                                    gcollect.Add(new MElement
                                                    {
                                                        ElementID = o,
                                                        ElementRecord = total.ToString(),
                                                        ElementDate = DateTime.Now.AddMonths(-1)
                                                    });
                                                }
                                            });
                                        }
                                    }
                                }
                                else
                                {
                                    x.Value.ForEach((o) =>
                                    {
                                        if (gcollect.FirstOrDefault(z => z.ElementID == o) == null)
                                        {
                                            var glist = gcollect.FirstOrDefault(z => z.ElementID == x.Key);
                                            if (glist != null)
                                            {
                                                gcollect.Add(new MElement
                                                {
                                                    ElementID = o,
                                                    ElementRecord = glist.ElementRecord,
                                                    ElementDate = DateTime.Now.AddMonths(-1)
                                                });
                                            }
                                        }
                                    });
                                }
                            }
                        }
                        /// 再次檢查&合併相同數值之要素
                        foreach (var x in gduplicate)
                        {
                            var glists = gcollect.FirstOrDefault(o => o.ElementID == x.Key);
                            if (glists != null)
                                continue;
                            x.Value.ForEach((o) =>
                            {
                                if (!o.Contains("+") && gcollect.FirstOrDefault(z => z.ElementID == o) == null)
                                {
                                    var glist = gcollect.FirstOrDefault(z => z.ElementID == x.Key);
                                    if (glist != null)
                                    {
                                        gcollect.Add(new MElement
                                        {
                                            ElementID = o,
                                            ElementRecord = glist.ElementRecord,
                                            ElementDate = DateTime.Now.AddMonths(-1)
                                        });
                                    }
                                }
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("匯入成功 : " + gcollect.Count());

            if (duplicate.Count > 0)
            {
                TxtBox1.Text += Environment.NewLine + "重複資料清單 : " + string.Join(",", duplicate) + Environment.NewLine;
            }
            try
            {
                if (gcollect.Count > 0)
                {
                    TxtBox1.Text += Environment.NewLine + string.Format("指標收回數量 : {0}/{1} ({2}%)", gcollect.Count, gdata.Count, gcollect.Count * 100 / gdata.Count) +
                        Environment.NewLine;
                    //gcollect.ForEach((o) => { TxtBox1.Text += string.Join(" , ", o.ElementID); });
                    gcollect.Sort((x, y) => { return x.ElementID.CompareTo(y.ElementID); });

                    string fpath = Environment.CurrentDirectory + @"\要素備份";
                    if (!Directory.Exists(fpath))
                    {
                        Directory.CreateDirectory(fpath);
                    }
                    string fname = @"\指標收集存檔(月份)" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx";
                    string fname2 = @"\指標收集存檔總檔.xlsx";
                    using (SLDocument nsl = new SLDocument())
                    {

                        nsl.SetColumnWidth(1, 15);
                        nsl.SetColumnWidth(2, 15);
                        nsl.SetCellValue(1, 1, "指標要素");
                        nsl.SetCellValue(1, 2, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                        for (int i = 0; i < gcollect.Count; i++)
                        {
                            nsl.SetCellValue(i + 2, 1, gcollect[i].ElementID);
                            nsl.SetCellValue(i + 2, 2, Convert.ToDouble(gcollect[i].ElementRecord));
                        }

                        nsl.SaveAs(fpath + fname);
                    }
                    if (!System.IO.File.Exists(fpath + fname))
                        return;

                    if (!System.IO.File.Exists(fpath + fname2))
                    {
                        using (SLDocument sl = new SLDocument(fpath + fname))
                        {
                            sl.SaveAs(fpath + fname2);
                        }
                    }
                    else
                    {
                        using (SLDocument sl = new SLDocument(fpath + fname2))
                        {
                            if (!DateTime.TryParse(sl.GetCellValueAsString(1, 2), out DateTime dts))
                            {
                                sl.CloseWithoutSaving();
                                return;
                            }
                            if ((dts.Month == DateTime.Now.Month && dts.Year == DateTime.Now.Year) || dts >= DateTime.Now)
                                return;
                            if (dts.Month != DateTime.Now.AddMonths(-1).Month &&
                                dts.Year != DateTime.Now.AddMonths(-1).Year)
                            {
                                sl.CopyColumn(2, 100, 3);
                                sl.SetCellValue(1, 2, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                            }

                            foreach (var x in gcollect)
                            {
                                bool oldele = false;
                                for (int j = 0; j < 500; j++)
                                {
                                    if (string.IsNullOrWhiteSpace(sl.GetCellValueAsString(j + 2, 1)))
                                        break;
                                    if (sl.GetCellValueAsString(j + 2, 1) == x.ElementID)
                                    {
                                        sl.SetCellValue(j + 2, 2, Convert.ToDouble(x.ElementRecord));
                                        oldele = true;
                                        break;
                                    }
                                }
                                if (!oldele)
                                {
                                    SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();
                                    int slrows = wsstats.EndRowIndex;

                                    sl.SetCellValue(slrows + 1, 1, x.ElementID);
                                    sl.SetCellValue(slrows + 1, 2, Convert.ToDouble(x.ElementRecord));
                                }
                            }

                            sl.SaveAs(fpath + fname2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void BT_TO_EXPORT_CLINIC(object sender, RoutedEventArgs e)
        {
            if (gdata.Count <= 0)
            {
                MessageBox.Show("尚未匯入指標清單");
                return;
            }
            if (!Directory.Exists(System.Environment.CurrentDirectory + @"\資料收集"))
            {
                Directory.CreateDirectory(System.Environment.CurrentDirectory + @"\資料收集");
            }
            ///
            /// 匯出時濾掉名稱不同但相同數值的要素
            ///
            var newda = gdata;
            newda.Sort((x, y) => { return x.ElementName.CompareTo(y.ElementName); });
            var newdata = newda.GroupBy(o => o.Depart)
                    .ToDictionary(o => o.Key, o => o.ToList());

            var unitcounts = gdata.Where(o => !SameEle.Contains(o.ElementID))
                .GroupBy(o => o.Depart)
                .ToDictionary(o => o.Key, o => o.ToList().Count);
            TxtBox1.Text += Environment.NewLine + "指標收集單位數 : " + unitcounts.Count +
                Environment.NewLine + string.Join(",", unitcounts) + Environment.NewLine;
            try
            {
                foreach (var x in newdata)
                {
                    SLDocument nsl = new SLDocument();
                    nsl.DocumentProperties.Creator = "TTMHH's QMC";
                    nsl.DocumentProperties.ContentStatus = "Measurement";
                    nsl.DocumentProperties.Title = "Measurement File";
                    nsl.DocumentProperties.Description = "Measurement File";
                    SLStyle style;
                    nsl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "指標提報");
                    style = nsl.CreateStyle();
                    style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                    style.Alignment.Vertical = VerticalAlignmentValues.Center;
                    style.Alignment.WrapText = true;
                    style.Font.FontSize = 12;
                    nsl.SetColumnWidth(1, 15);
                    nsl.SetColumnWidth(2, 10);
                    nsl.SetColumnWidth(3, 25);
                    nsl.SetColumnWidth(4, 45);
                    nsl.SetColumnWidth(5, 10);
                    nsl.SetColumnStyle(1, style);
                    nsl.SetColumnStyle(2, style);
                    nsl.SetColumnStyle(3, style);
                    nsl.SetColumnStyle(4, style);

                    nsl.SetCellValue(1, 1, "指標群組");
                    nsl.SetCellValue(1, 2, "監測單位");
                    nsl.SetCellValue(1, 3, "指標要素");
                    nsl.SetCellValue(1, 4, "指標(要素)名稱");
                    for (int i = 0; i < 6; i++)
                    {
                        nsl.SetColumnStyle(i + 5, style);
                        nsl.SetCellValue(1, i + 5, DateTime.Now.AddMonths(-(i + 1)).ToString("yyyy/MM"));
                    }
                    int index = 2;
                    style.Protection.Locked = false;
                    style.Font.FontColor = System.Drawing.Color.DarkMagenta;
                    style.Font.FontSize = 13;
                    style.Font.Bold = true;
                    style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightCyan, System.Drawing.Color.White);
                    style.SetRightBorder(BorderStyleValues.Thin, System.Drawing.Color.LightGray);
                    style.SetLeftBorder(BorderStyleValues.Thin, System.Drawing.Color.LightGray);
                    style.SetTopBorder(BorderStyleValues.Thin, System.Drawing.Color.LightGray);
                    style.SetBottomBorder(BorderStyleValues.Thin, System.Drawing.Color.LightGray);
                    for (int i = 0; i < x.Value.Count; i++)
                    {
                        if (x.Value[i].Complex)
                            continue;
                        if (SameEle.Count > 0 && SameEle.Contains(x.Value[i].ElementID))
                        {
                            continue;
                        }
                        nsl.SetCellValue(index, 1, x.Value[i].Group);
                        nsl.SetCellValue(index, 2, x.Value[i].Depart);
                        nsl.SetCellValue(index, 3, x.Value[i].ElementID);
                        nsl.SetCellValue(index, 4, x.Value[i].ElementName);

                        nsl.SetCellStyle(index, 5, style);

                        if (gbackups.Count > 0 && gbackups.ContainsKey(x.Value[i].ElementID))
                        {
                            for (int j = 2; j < 7; j++)
                            {
                                var data = gbackups[x.Value[i].ElementID].Find(o => o.ElementDate.Month == DateTime.Now.AddMonths(-j).Month
                                && o.ElementDate.Year == DateTime.Now.AddMonths(-j).Year);
                                if (data == null)
                                    continue;

                                nsl.SetCellValueNumeric(index, j + 4, data.ElementRecord);
                            }
                        }
                        index++;
                    }
                    SLSheetProtection sp = new SLSheetProtection();
                    sp.AllowInsertRows = false;
                    sp.AllowInsertColumns = false;
                    sp.AllowFormatCells = true;
                    sp.AllowDeleteColumns = false;
                    sp.AllowDeleteRows = false;
                    sp.AllowSelectUnlockedCells = true;
                    sp.AllowSelectLockedCells = false;
                    nsl.ProtectWorksheet(sp);

                    string Refile = System.Environment.CurrentDirectory + @"\資料收集\" + x.Key + ".xlsx";
                    nsl.SaveAs(Refile);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageBox.Show("轉出成功 : " + newdata.Count);
        }

        private void BT_IMPORT_OLDDATA(object sender, RoutedEventArgs e)
        {
            if (gmeasure.Count <= 0)
            {
                MessageBox.Show("請先匯入指標資料");
                return;
            }
            /*
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = Environment.CurrentDirectory;
            dlg.Title = "選取資料檔";
            dlg.Filter = "xlsx files (*.*)|*.xlsx";
            if (dlg.ShowDialog() != true)
                return;
            if (!System.IO.File.Exists(dlg.FileName))
                return;
                */
            string folderName = System.Environment.CurrentDirectory + @"\過往提報資料\THIS";
            try
            {
                foreach (var finame in System.IO.Directory.GetFileSystemEntries(folderName))
                {
                    if (System.IO.Path.GetExtension(finame) != ".xlsx")
                        continue;
                    using (SLDocument sl = new SLDocument(finame))
                    {
                        for (int i = 0; i < 500; i++)
                        {
                            if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                                break;
                            if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                                continue;
                            if (!DateTime.TryParse(((sl.GetCellValueAsInt32(i + 2, 2) + 1911).ToString() + "/" + sl.GetCellValueAsString(i + 2, 3)), out DateTime dts))
                                continue;
                            var measureid = gmeasure.Where(o => o.MeasureID == sl.GetCellValueAsString(i + 2, 4)).FirstOrDefault();
                            if (measureid == null)
                                continue;

                            try
                            {
                                if (measureid.Numerator != "1")
                                {
                                    List<MElement> lmeNume = new List<MElement>();
                                    MElement dataNume = new MElement
                                    {
                                        ElementID = measureid.Numerator,
                                        ElementRecord = sl.GetCellValueAsString(i + 2, 5).Trim(),
                                        ElementDate = dts
                                    };
                                    if (!gbackups.ContainsKey(measureid.Numerator))
                                    {

                                        lmeNume.Add(dataNume);
                                        gbackups[measureid.Numerator] = lmeNume;
                                    }
                                    else
                                    {
                                        var meNume = gbackups[measureid.Numerator].FirstOrDefault(o => o.ElementDate == dts);
                                        if (meNume == null)
                                        {
                                            gbackups[measureid.Numerator].Add(dataNume);
                                        }
                                    }
                                }
                                if (measureid.Denominator != "1")
                                {
                                    List<MElement> lmeDeno = new List<MElement>();
                                    MElement dataDeno = new MElement
                                    {
                                        ElementID = measureid.Denominator,
                                        ElementRecord = sl.GetCellValueAsString(i + 2, 6).Trim(),
                                        ElementDate = dts
                                    };
                                    if (!gbackups.ContainsKey(measureid.Denominator))
                                    {

                                        lmeDeno.Add(dataDeno);
                                        gbackups[measureid.Denominator] = lmeDeno;
                                    }
                                    else
                                    {
                                        var meDeno = gbackups[measureid.Denominator].FirstOrDefault(o => o.ElementDate == dts);
                                        if (meDeno == null)
                                        {
                                            gbackups[measureid.Denominator].Add(dataDeno);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                            }
                        }

                        sl.CloseWithoutSaving();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageBox.Show("匯入過往THIS資料成功");
        }

        private void BT_EXPORT_ELEMENT(object sender, RoutedEventArgs e)
        {
            if (gbackups.Count <= 0)
                return;

            if (gcollect.Count > 0)
            {
                foreach (var x in gcollect)
                {
                    if (gbackups.ContainsKey(x.ElementID))
                    {
                        var dataex = gbackups[x.ElementID].FirstOrDefault(o => o.ElementID == x.ElementID && o.ElementDate == x.ElementDate);
                        if (dataex != null)
                            gbackups[x.ElementID].Remove(dataex);
                        gbackups[x.ElementID].Add(x);
                    }
                }
            }

            var sortbacks = gbackups.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);

            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標收集總存檔" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx";
            //string fname2 = @"\指標收集存檔總檔.xlsx";
            try
            {
                using (SLDocument sl = new SLDocument())
                {
                    sl.SetColumnWidth(1, 15);
                    //sl.SetColumnWidth(2, 15);
                    sl.SetCellValue(1, 1, "指標要素");
                    SLStyle style = sl.CreateStyle();
                    for (int i = 0; i < 12 + 6; i++)
                        sl.SetCellValue(1, i + 2, DateTime.Now.AddMonths(-1 - i).ToString("yyyy/MM"));
                    int index = 0;
                    foreach (var x in sortbacks)
                    {
                        sl.SetCellValue(index + 2, 1, x.Key);

                        foreach (var y in x.Value)
                        {
                            for (int i = 0; i < 12 + 6; i++)
                            {
                                if (y.ElementDate.Year == DateTime.Now.AddMonths(-1 - i).Year
                                    && y.ElementDate.Month == DateTime.Now.AddMonths(-1 - i).Month && Double.TryParse(y.ElementRecord, out double num))
                                {
                                    sl.SetCellValue(index + 2, i + 2, num);
                                }
                            }
                        }

                        index++;
                    }
                    
                    sl.SaveAs(fpath + fname);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageBox.Show("指標匯出結束");
        }

        private void BT_EXPORT_MEASURE(object sender, RoutedEventArgs e)
        {
            if (gmeasure.Count <= 0)
                return;
            if (gbackups.Count <= 0)
                return;
            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標數據總資料" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx";

            var sortbacks = gbackups.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value.OrderByDescending(o => o.ElementRecord).ToList());
            try
            {
                using (SLDocument sl = new SLDocument())
                {
                    SLStyle style;
                    style = sl.CreateStyle();
                    style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                    style.Alignment.Vertical = VerticalAlignmentValues.Center;
                    style.Alignment.WrapText = true;
                    style.Font.FontSize = 12;
                    sl.SetColumnWidth(1, 15);
                    sl.SetColumnWidth(2, 10);
                    sl.SetColumnWidth(3, 40);
                    sl.SetColumnStyle(1, style);
                    sl.SetColumnStyle(2, style);
                    sl.SetColumnStyle(3, style);

                    sl.SetCellValue(1, 1, "指標群組");
                    sl.SetCellValue(1, 2, "指標代號");
                    sl.SetCellValue(1, 3, "指標名稱");
                    for (int i = 0; i < 12 + 6; i++)
                        sl.SetCellValue(1, i + 4, DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM"));

                    int index = 2;
                    foreach (var x in gmeasure)
                    {
                        sl.SetCellValue(index, 1, x.Group);
                        sl.SetCellValue(index, 2, x.MeasureID);
                        sl.SetCellValue(index, 3, x.MeasureName);
                        sl.MergeWorksheetCells(index, 1, index + 2, 1);
                        sl.MergeWorksheetCells(index, 2, index + 2, 2);
                        sl.MergeWorksheetCells(index, 3, index + 2, 3);
                        int Destatus = 0;
                        int Nustatus = 0;

                        List<List<MElement>> DenosPlus = new List<List<MElement>>();
                        List<List<MElement>> NumePlus = new List<List<MElement>>();

                        var Numes = sortbacks.FirstOrDefault(o => o.Key == x.Numerator).Value;
                        var Denos = sortbacks.FirstOrDefault(o => o.Key == x.Denominator).Value;

                        if (x.Denominator.Contains("+"))
                        {
                            Destatus = 1;
                            var elements = x.Denominator.Split('+').ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele.Trim()).Value;
                                    if (em != null)
                                        DenosPlus.Add(em);
                                }
                            }
                        }
                        else if (x.Denominator.Contains(".") && x.Denominator.Contains("-"))
                        {
                            Destatus = 2;
                            var elements = x.Denominator.Split('-').ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele.Trim()).Value;
                                    if (em != null)
                                        DenosPlus.Add(em);
                                }
                            }
                        }
                        if (x.Numerator.Contains("+"))
                        {
                            Nustatus = 1;
                            var elements = x.Numerator.Split('+').ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele.Trim()).Value;
                                    if (em != null)
                                        NumePlus.Add(em);
                                }
                            }
                        }
                        else if (x.Numerator.Contains(".") && x.Numerator.Contains("-"))
                        {
                            Nustatus = 2;
                            var elements = x.Numerator.Split('-').ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele.Trim()).Value;
                                    if (em != null)
                                        NumePlus.Add(em);
                                }
                            }
                        }
                        for (int i = 0; i < 12 + 6; i++)
                        {
                            if (x.Numerator == "1")
                                sl.SetCellValue(index + 1, i + 4, 1);
                            else if (Numes != null)
                            {
                                var nume = Numes.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                && o.ElementDate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                if (nume != null)
                                {
                                    if (Double.TryParse(nume.ElementRecord, out double numok))
                                        sl.SetCellValue(index + 1, i + 4, numok);
                                }
                            }
                            else if (NumePlus.Count > 0)
                            {
                                try
                                {
                                    Double deno = 0;
                                    foreach (var ele in NumePlus)
                                    {
                                        var de = ele.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                    && o.ElementDate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                        if (de == null)
                                        {
                                            break;
                                        }
                                        if (!Double.TryParse(de.ElementRecord, out double num))
                                            break;
                                        if (Nustatus == 1)
                                        {
                                            deno += num;
                                        }
                                        else if (Nustatus == 2)
                                        {
                                            if (deno == 0)
                                                deno = num;
                                            else
                                                deno -= num;
                                        }
                                    }
                                    if (deno > 0)
                                        sl.SetCellValue(index + 2, i + 4, deno);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                            }
                            if (x.Denominator == "1")
                                sl.SetCellValue(index + 2, i + 4, "NA");
                            else if (Denos != null)
                            {
                                var deno = Denos.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                && o.ElementDate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                if (deno != null)
                                {
                                    if (Double.TryParse(deno.ElementRecord, out double numok))
                                        sl.SetCellValue(index + 2, i + 4, numok);
                                }
                            }
                            else if (DenosPlus.Count > 0)
                            {
                                try
                                {
                                    Double deno = 0;
                                    foreach (var ele in DenosPlus)
                                    {
                                        var de = ele.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                    && o.ElementDate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                        if (de == null)
                                        {
                                            break;
                                        }
                                        if (!Double.TryParse(de.ElementRecord, out double num))
                                            break;
                                        if (Destatus == 1)
                                        {
                                            deno += num;
                                        }
                                        else if (Destatus == 2)
                                        {
                                            if (deno == 0)
                                                deno = num;
                                            else
                                                deno -= num;
                                        }
                                    }
                                    if (deno > 0)
                                        sl.SetCellValue(index + 2, i + 4, deno);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                            }
                            //將分子、分母數據加回 gmeasure
                            if (!x.Records.ContainsKey(DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM")))
                                x.Records.Add(DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM"),
                                new List<double>() { sl.GetCellValueAsDouble(index + 1, i + 4), sl.GetCellValueAsDouble(index + 2, i + 4) });

                            if (!string.IsNullOrEmpty(sl.GetCellValueAsString(index + 2, i + 4)))
                            {
                                if (Double.TryParse(sl.GetCellValueAsString(index + 1, i + 4), out double nu)
                                    && Double.TryParse(sl.GetCellValueAsString(index + 2, i + 4) == "NA" ? "1" : sl.GetCellValueAsString(index + 2, i + 4), out double de))
                                {
                                    if (nu == 0)
                                        sl.SetCellValue(index, i + 4, 0);
                                    else if (de != 0)
                                        sl.SetCellValue(index, i + 4, nu / de);
                                }
                            }
                        }
                        index += 3;
                    }
                    
                    sl.SaveAs(fpath + fname);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageBox.Show("指標匯出結束");
        }

        private void BT_EXPORT_CHART(object sender, RoutedEventArgs e)
        {
            if (gmeasure.Count < 0)
                return;
            SPCtype spctype;
            if (Combx3.SelectedIndex < 0 || Combx2.SelectedIndex < 0)
                return;
            else if (Combx3.SelectedIndex == 0)
            {
                spctype = SPCtype.P;
            }
            else
                spctype = (SPCtype)Combx3.SelectedIndex;

            var selmeasure = Combx2.SelectedValue.ToString().Split(':').FirstOrDefault();
            var smeasure = gmeasure.FirstOrDefault(o => o.MeasureID == selmeasure);
            if (smeasure == null || smeasure.Records.Count <= 0)
                return;

            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            try
            {
                List<string> alphabet = Enumerable.Range(0, 26).Select(i => Convert.ToChar('A' + i).ToString()).ToList();
                using (SLDocument sl = new SLDocument())
                {
                    sl.SetCellValue(2, 2, Combx2.SelectedValue.ToString());
                    sl.SetCellValue(3, 2, "平均值");
                    sl.SetCellValue(4, 2, "管制圖上限 (2α)");
                    sl.SetCellValue(5, 2, "管制圖下限 (2α)");
                    sl.SetCellValue(6, 2, "管制圖上限 (3α)");
                    sl.SetCellValue(7, 2, "管制圖下限 (3α)");
                    sl.SetCellValue(8, 2, "分子" + (char)10 + smeasure.NumeratorName);
                    sl.SetCellValue(9, 2, "分母" + (char)10 + smeasure.DenominatorName);
                    //取得數值資料
                    for (int i = 0; i < 12; i++)
                    {
                        var srecord = smeasure.Records.Where(o => DateTime.TryParse(o.Key, out DateTime sda)
                        && sda.Year == DateTime.Now.AddMonths(-12 - 1 + i).Year
                        && sda.Month == DateTime.Now.AddMonths(-12 - 1 + i).Month).ToDictionary(o => o.Key, o => o.Value);
                        if (srecord == null || srecord.Count <= 0)
                            continue;

                        sl.SetCellValue(1, 3 + i, srecord.Keys.FirstOrDefault());
                        sl.SetCellValue(8, 3 + i, srecord.FirstOrDefault().Value[0]);
                        sl.SetCellValue(9, 3 + i, srecord.FirstOrDefault().Value[1]);
                        //sl.SetCellValue(1, 3 + i, sl.GetCellValueAsString(1, 3 + 12 - i));
                        //sl.SetCellValue(8, 3 + i, sl.GetCellValueAsDouble(3, 3 + 12 - i));
                        //sl.SetCellValue(9, 3 + i, sl.GetCellValueAsDouble(4, 3 + 12 - i));
                    }
                    if (spctype == SPCtype.C)
                    {
                        sl.SetCellValue(2, 4 + 12 + 2, "c");
                        sl.SetCellValue(3, 4 + 12 + 2, "=SUM(C8:N8) / 12");
                        sl.SetCellValue(2, 5 + 12 + 2, "n");
                        sl.SetCellValue(3, 5 + 12 + 2, 12);
                    }
                    else
                    {
                        sl.SetCellValue(2, 4 + 12 + 2, "u or p");
                        sl.SetCellValue(3, 4 + 12 + 2, "=SUM(C8:N8) / SUM(C9:N9)");
                        sl.SetCellValue(2, 5 + 12 + 2, "n");
                        sl.SetCellValue(3, 5 + 12 + 2, "=AVERAGE(C9:N9)");
                    }

                    for (int i = 0; i < 12; i++)
                    {
                        string sigma = string.Empty;
                        string amplify = string.Empty;
                        string average = "R3";
                        string measure = string.Empty;
                        if (spctype == SPCtype.P)
                        {
                            amplify = " * 1000";
                            sigma = string.Format("SQRT(R3 * (1 - R3) / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        else if (spctype == SPCtype.U)
                        {
                            amplify = string.Empty;
                            sigma = string.Format("SQRT(R3 / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        else if (spctype == SPCtype.C)
                        {
                            amplify = string.Empty;
                            sigma = "SQRT(R3)";
                            measure = string.Format("=({0}8)", alphabet[2 + i]);
                        }
                        else if (spctype == SPCtype.nP)
                        {
                            average = "S3 * R3";
                            amplify = string.Empty;
                            sigma = "SQRT(S3 * R3 * (1 - R3))";
                            measure = string.Format("=({0}8)", alphabet[2 + i]);
                        }
                        else //暫時借用 U 圖
                        {
                            amplify = string.Empty;
                            sigma = string.Format("SQRT(R3 / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        if (sl.GetCellValueAsDouble(9, 3 + i) == 0)
                            sl.SetCellValue(2, 3 + i, 0);
                        else
                            sl.SetCellValue(2, 3 + i, measure);
                        sl.SetCellValue(3, 3 + i, string.Format("=({0}){1}", average, amplify));
                        sl.SetCellValue(4, 3 + i, string.Format("=({0} + ({1} * 2)){2}", average, sigma, amplify));
                        sl.SetCellValue(5, 3 + i, string.Format("=({0} - ({1} * 2)){2}", average, sigma, amplify));
                        sl.SetCellValue(6, 3 + i, string.Format("=({0} + ({1} * 3)){2}", average, sigma, amplify));
                        sl.SetCellValue(7, 3 + i, string.Format("=({0} - ({1} * 3)){2}", average, sigma, amplify));
                    }

                    sl.SetColumnWidth(2, 25);
                    SLStyle style;
                    style = sl.CreateStyle();
                    style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                    style.Alignment.Vertical = VerticalAlignmentValues.Center;
                    style.Alignment.WrapText = true;
                    style.Font.FontSize = 12;
                    sl.SetColumnStyle(2, style);
                    sl.SetRowStyle(1, style);

                    double fChartHeight = 25;
                    double fChartWidth = 13;
                    SLChart chart;
                    chart = sl.CreateChart("B1", "N7");
                    chart.SetChartType(SLLineChartType.Line);
                    chart.SetChartStyle(SLChartStyle.Style5);
                    chart.SetChartPosition(11, 1, 11 + fChartHeight, 1 + fChartWidth);

                    chart.PrimaryTextAxis.TickLabelPosition = DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues.Low;

                    SLDataSeriesOptions dso;
                    dso = chart.GetDataSeriesOptions(3);
                    dso.Line.DashType = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot;
                    chart.SetDataSeriesOptions(3, dso);
                    chart.SetDataSeriesOptions(4, dso);
                    dso.Line.DashType = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.DashDot;
                    chart.SetDataSeriesOptions(5, dso);
                    chart.SetDataSeriesOptions(6, dso);

                    dso = chart.GetDataSeriesOptions(1);
                    dso.Marker.Symbol = DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues.Circle;
                    dso.Line.SetSolidLine(System.Drawing.Color.Chocolate, 0);
                    chart.SetDataSeriesOptions(1, dso);
                    sl.InsertChart(chart);
                    sl.SaveAs(fpath + @"\(" + Combx3.SelectedValue.ToString() + ")SPCChart-指標數據總資料 (" + smeasure.MeasureName + ") " + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            /*
            string fname = @"\指標數據總資料" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx";
            if (!System.IO.File.Exists(fpath+fname))
                return;

            try
            {
                List<string> alphabet = Enumerable.Range(0, 26).Select(i => Convert.ToChar('A' + i).ToString()).ToList();
                using (SLDocument sl = new SLDocument(fpath + fname))
                {
                    SLDocument sl2 = new SLDocument();
                    sl2.SetCellValue(2, 2, sl.GetCellValueAsString(2, 2) + sl.GetCellValueAsString(2, 3) + "(‰)");
                    sl2.SetCellValue(3, 2, "平均值");
                    sl2.SetCellValue(4, 2, "管制圖上限 (2α)");
                    sl2.SetCellValue(5, 2, "管制圖下限 (2α)");
                    sl2.SetCellValue(6, 2, "管制圖上限 (3α)");
                    sl2.SetCellValue(7, 2, "管制圖下限 (3α)");
                    sl2.SetCellValue(8, 2, "分子(死亡人數(含病危自動出院))");
                    sl2.SetCellValue(9, 2, "分母(死亡人數(含病危自動出院)+出院人次 (不含死亡及病危自動出院))");
                    //double datotal = 0;
                    //List<double> datalist = new List<double>();
                    for (int i = 0; i < 12; i++)
                    {
                        sl2.SetCellValue(1, 3 + i, sl.GetCellValueAsString(1, 3 + 12 - i));
                        sl2.SetCellValue(8, 3 + i, sl.GetCellValueAsDouble(3, 3 + 12 - i));
                        sl2.SetCellValue(9, 3 + i, sl.GetCellValueAsDouble(4, 3 + 12 - i));
                    }
                    //sl2.SetCellValue(2, 3 + 12 + 2, "α");
                    sl2.SetCellValue(2, 4 + 12 + 2, "u or p");
                    sl2.SetCellValue(3, 4 + 12 + 2, "=SUM(C8:N8) / SUM(C9:N9)");
                    sl2.SetCellValue(2, 5 + 12 + 2, "n");
                    sl2.SetCellValue(3, 5 + 12 + 2, "=AVERAGE(C9:N9)");
                    //var datastatic = StSummary(datalist);
                    for (int i = 0; i < 12; i++)
                    {
                        //sl2.SetCellValue(4, 3 + i, datotal / 6);
                        //sl2.SetCellValue(5, 3 + i, sl2.GetCellValueAsDouble(4, 3) + sl2.GetCellValueAsDouble(4, 3 + 7) * 2);
                        //sl2.SetCellValue(6, 3 + i, sl2.GetCellValueAsDouble(4, 3) - sl2.GetCellValueAsDouble(4, 3 + 7) * 2);
                        //sl2.SetCellValue(5, 3 + i, datastatic["Mean"] + datastatic["StandardDeviation"] * 2);
                        //sl2.SetCellValue(6, 3 + i, datastatic["Mean"] - datastatic["StandardDeviation"] * 2);
                        //sl2.SetCellValue(2, 3 + i, sl.GetCellValueAsDouble(2, 3 + 12 - i) * 10);
                        //sl2.SetCellValue(2, 3 + i, sl2.GetCellValueAsDouble(9, 3 + i) == 0 ? "0" :((sl2.GetCellValueAsDouble(8, 3 + i) / sl2.GetCellValueAsDouble(9, 3 + i)) * 100).ToString());
                        //sl2.SetCellValue(3, 3 + i, "=AVERAGE(C2:N2)");
                        string sigma = string.Empty;
                        string amplify = string.Empty;
                        string average = "R3";
                        string measure = string.Empty;
                        if (spctype == SPCtype.P)
                        {
                            amplify = " * 1000";
                            sigma = string.Format("SQRT(R3 * (1 - R3) / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        else if (spctype == SPCtype.U)
                        {
                            amplify = string.Empty;
                            sigma = string.Format("SQRT(R3 / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        else if (spctype == SPCtype.C)
                        {
                            amplify = string.Empty;
                            sigma = "SQRT(R3)";
                            measure = string.Format("=({0}8)", alphabet[2 + i]);
                        }
                        else if (spctype == SPCtype.nP)
                        {
                            average = "S3 * R3";
                            amplify = string.Empty;
                            sigma = "SQRT(S3 * R3 * (1 - R3))";
                            measure = string.Format("=({0}8)", alphabet[2 + i]);
                        }
                        if (sl2.GetCellValueAsDouble(9, 3 + i) == 0)
                            sl2.SetCellValue(2, 3 + i, 0);
                        else
                            sl2.SetCellValue(2, 3 + i, measure);
                        sl2.SetCellValue(3, 3 + i, string.Format("=({0}){1}", average, amplify));
                        sl2.SetCellValue(4, 3 + i, string.Format("=({0} + ({1} * 2)){2}", average, sigma, amplify));
                        sl2.SetCellValue(5, 3 + i, string.Format("=({0} - ({1} * 2)){2}", average, sigma, amplify));
                        sl2.SetCellValue(6, 3 + i, string.Format("=({0} + ({1} * 3)){2}", average, sigma, amplify));
                        sl2.SetCellValue(7, 3 + i, string.Format("=({0} - ({1} * 3)){2}", average, sigma, amplify));
                    }

                    sl2.SetColumnWidth(2, 25);
                    SLStyle style;
                    style = sl2.CreateStyle();
                    style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                    style.Alignment.Vertical = VerticalAlignmentValues.Center;
                    style.Alignment.WrapText = true;
                    style.Font.FontSize = 12;
                    sl2.SetColumnStyle(2, style);
                    sl2.SetRowStyle(1, style);

                    double fChartHeight = 20;
                    double fChartWidth = 10;
                    SLChart chart;
                    chart = sl2.CreateChart("B1", "N7");
                    chart.SetChartType(SLLineChartType.Line);
                    chart.SetChartStyle(SLChartStyle.Style5);
                    chart.SetChartPosition(11, 2, 11 + fChartHeight, 2 + fChartWidth);

                    chart.PrimaryTextAxis.TickLabelPosition = DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues.Low;

                    SLDataSeriesOptions dso;
                    dso = chart.GetDataSeriesOptions(3);
                    dso.Line.DashType = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot;
                    chart.SetDataSeriesOptions(3, dso);
                    chart.SetDataSeriesOptions(4, dso);
                    dso.Line.DashType = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.DashDot;
                    chart.SetDataSeriesOptions(5, dso);
                    chart.SetDataSeriesOptions(6, dso);

                    dso = chart.GetDataSeriesOptions(1);
                    dso.Marker.Symbol = DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues.Circle;
                    dso.Line.SetSolidLine(System.Drawing.Color.Chocolate, 0);
                    chart.SetDataSeriesOptions(1, dso);
                    sl2.InsertChart(chart);
                    sl2.SaveAs(fpath + @"\Chart-指標數據總資料" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            */
            MessageBox.Show("匯出圖表成功");
        }

        private void BT_EXPORT_CUSTOM(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = Environment.CurrentDirectory;
            dlg.Title = "選取自訂檔案";
            dlg.Filter = "MES files (*.mes)|*.mes";
            if (dlg.ShowDialog() == true)
            {
                LoadCustom(dlg.FileName);
            }
        }
        public void LoadCustom(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return;
            string[] datas = File.ReadAllLines(fname);
            string data = datas.FirstOrDefault(x => x.Substring(0, 1) != "#");
            List<string> customs = data.Split(';').ToList();
            if (customs.Count <= 0)
                return;
            var incus = customs.ConvertAll(s => Int32.Parse(s)).ToList();
            List<MElement> sdata = new List<MElement>();
            List<MMeasure> smeasure = new List<MMeasure>();
            string gtype;
            switch (incus[0])
            {
                case 0:
                    gtype = "TCPI";
                    break;
                case 1:
                    gtype = "THIS";
                    break;
                case 2:
                    gtype = "HACMI";
                    break;
                case 3:
                    gtype = "";
                    break;
                default:
                    gtype = "TCPI";
                    break;
            };
            if (!string.IsNullOrEmpty(gtype))
                sdata = gdata.Where(x => x.Group == gtype).ToList();
            else
                sdata = gdata;
    }
        private enum SPCtype
        {
            U = 1,
            C = 2,
            P = 3,
            I = 4,
            nP = 5,
            Xbar_S = 6,
            Xbar_R = 7
        }
        private enum IndicatorGroup
        {
            TCPI,
            THIS,
            HACMI,
            ALL
        }
        private enum CustomType
        {
            E_Group,
            E_Depart,
            E_ELEID,
            E_ELEName,
            E_Record,
            E_Frequency,
            E_User,
            E_MeasureID,
            E_MeasureName,
            E_Numerator,
            E_NumeratorName,
            E_Denominator,
            E_DenominatorName,
            E_Threshold,
            E_Date_1,
            E_Date_2,
            E_Date_3,
            E_Date_4,
            E_Date_5,
            E_Date_6,
            E_Date_7,
            E_Date_8,
            E_Date_9,
            E_Date_10,
            E_Date_11,
            E_Date_12,
        }

        private void Combx1_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (gmeasure.Count <= 0)
                return;

            if (Combx1.SelectedIndex == 0)
            {
                Combx2.ItemsSource = gmeasure.Select(o => o.MeasureID + ":" + o.MeasureName).ToList();
            }
            else
            {
                Combx2.ItemsSource = gmeasure.Where(o => o.Group == Combx1.SelectedValue.ToString()).Select(o => o.MeasureID + ":" + o.MeasureName).ToList();
            }
        }
    }
}
