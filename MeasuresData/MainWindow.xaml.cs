using System;
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
    #region CLASS

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

        public static MMeasure operator +(MMeasure a, MMeasure b)
        {
            MMeasure temp = new MMeasure();
            if (a.Records.Count <= 0 || b.Records.Count <= 0)
                return temp;
            List<double> rds = new List<double>();
            foreach (var x in a.Records)
            {
                var rd2 = b.Records.FirstOrDefault(o => o.Key == x.Key).Value;
                if (rd2.Count == 2 && x.Value.Count == 2)
                {
                    temp.Records[x.Key] = new List<double>() { x.Value[0] + rd2[0], x.Value[1] + rd2[1] };
                }
            }
            return temp;
        }
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
        public Dictionary<string, double> AllRecords { get; set; }
        public MElement()
        {
            AllRecords = new Dictionary<string, double>();
            ElementDate = DateTime.MinValue;
        }
    }
    public enum SPCtype
    {
        U = 1,
        C = 2,
        P = 3,
        I = 4,
        nP = 5,
        Xbar_S = 6,
        Xbar_R = 7
    }
    public class SPC
    {
        public string Name { get; set; }
        public SPCtype Type { get; set; }
        public Dictionary<string, List<double>> Records { get; set; }
        public SPC()
        {
            Records = new Dictionary<string, List<double>>();
            Type = SPCtype.P;
        }
        public SPC(Dictionary<string, List<double>> records, SPCtype type)
        {
            Type = type;
            if (records.Count > 0)
            {
                foreach (var x in records)
                {
                    sum_a += x.Value[0];
                    sum_b += x.Value[1];
                }
                if (sum_b == 0)
                    return;
                if (Type == SPCtype.C)
                {
                    u_p = sum_a / records.Count;
                    n_n = records.Count;
                }
                else
                {
                    u_p = sum_a / sum_b;
                    n_n = sum_b / records.Count;
                }
                avg = Type == SPCtype.nP ? u_p * n_n : u_p;
                Title = new List<string>();
                Measures = new List<double>();
                Average = new List<double>();
                UCL = new List<double>();
                LCL = new List<double>();
                UUCL = new List<double>();
                LLCL = new List<double>();
                foreach (var x in records)
                {
                    if (Type == SPCtype.P)
                    {
                        sigma = x.Value[1] == 0 ? 0 : Math.Sqrt(u_p * (1 - u_p) / x.Value[1]);
                        measure = x.Value[1] == 0 ? 0 : x.Value[0] / x.Value[1];
                    }
                    else if (Type == SPCtype.U)
                    {
                        sigma = x.Value[1] == 0 ? 0 : Math.Sqrt(u_p / x.Value[1]);
                        measure = x.Value[1] == 0 ? 0 : x.Value[0] / x.Value[1];
                    }
                    else if (Type == SPCtype.C)
                    {
                        sigma = Math.Sqrt(u_p);
                        measure = x.Value[0];
                    }
                    else if (Type == SPCtype.nP)
                    {
                        sigma = Math.Sqrt(n_n * u_p * (1 - u_p));
                        measure = x.Value[0];
                    }
                    else //暫時借用 U 圖
                    {
                        sigma = x.Value[1] == 0 ? 0 : Math.Sqrt(u_p / x.Value[1]);
                        measure = x.Value[1] == 0 ? 0 : x.Value[0] / x.Value[1];
                    }

                    Measures.Add(measure);
                    Average.Add(avg);
                    UCL.Add(avg + (sigma * 3));
                    LCL.Add(avg - (sigma * 3));
                    UUCL.Add(avg + (sigma * 2));
                    LLCL.Add(avg - (sigma * 2));
                    Title.Add(x.Key);
                }
            }
        }
        public List<string> Title { get; set; }
        public List<double> Measures { get; set; }
        public List<double> Average { get; set; }
        public List<double> UCL { get; set; }
        public List<double> LCL { get; set; }
        public List<double> UUCL { get; set; }
        public List<double> LLCL { get; set; }
        private double sum_a;
        private double sum_b;
        private double u_p;
        private double n_n;
        private double sigma;
        private double avg;
        private double measure;
        public double Avg
        {
            get
            {
                return Type == SPCtype.nP ? u_p * n_n : u_p;
            }
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
    #endregion

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
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            List<string> exmonth = new List<string>();
            for (int i = 0; i < 12 + 6; i++)
                exmonth.Add(DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM"));
            Combx4_Month.ItemsSource = Combx4_Month_ST.ItemsSource = Combx4_Month_Ed.ItemsSource = exmonth;
            Combx4_Month.SelectedIndex = 0;
            Combx4_Month_Ed.SelectedIndex = 1;
            Combx4_Month_ST.SelectedIndex = 12;
            /*
            List<string> assoicate = new List<string>() { "TCPI", "THIS", "評鑑持續", "P4P" };
            Combx5.ItemsSource = assoicate;
            Combx5.SelectedIndex = 0;
            */
            this.Cap.Title += " " + System.IO.File.GetLastWriteTime(this.GetType().Assembly.Location).ToString("yyy-MM-dd HH:mm");
        }
        #region Stastic
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
        #region PARAMETER
        /*
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
        */
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
        //原始資料庫之要素資料定義
        public List<MElement> GElements = new List<MElement>();
        //原始資料庫之指標資料定義
        public List<MMeasure> GMeasures = new List<MMeasure>();
        //收回單位填寫之要素數值
        public List<MElement> gcollect = new List<MElement>();
        //資料庫中過往要素數值
        public Dictionary<string, List<MElement>> GRecords = new Dictionary<string, List<MElement>>();
        //原始資料庫之重覆要素
        public Dictionary<string, List<string>> gduplicate = new Dictionary<string, List<string>>();
        public List<string> SameEle = new List<string>();

        //單位(群組合併)定義
        public Dictionary<string, List<string>> GMulgroups = new Dictionary<string, List<string>>();
        #endregion

        #region METHOD
        public void LoadFile(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return;
            GElements.Clear();
            GMeasures.Clear();
            gduplicate.Clear();
            SameEle.Clear();
            GRecords.Clear();
            try
            {
                SLDocument sl = new SLDocument(fname, "工作表1");

                for (int i = 0; i < 1000; i++)
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

                        if (gduplicate.ContainsKey(content) && gduplicate[content].Contains(data.ElementID))
                            continue;
                        if (!SameEle.Contains(content))
                            SameEle.Add(content);
                        if (data.Complex)
                            SameEle.Add(data.ElementID);

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

                    GElements.Add(data);
                }
                sl.CloseWithoutSaving();

                if (GElements.Count > 0)
                {
                    //MessageBox.Show("成功匯入要素數 : " + GElements.Count.ToString());
                    TxtBox1.Text += Environment.NewLine + "要素匯入數量 : " + GElements.Count + Environment.NewLine;
                    if (gduplicate.Count > 0)
                    {
                        TxtBox1.Text += Environment.NewLine + "相同意義要素組數量 : " + gduplicate.Count + Environment.NewLine;
                    }
                }
                else
                {
                    MessageBox.Show("匯入失敗");
                }

                sl = new SLDocument(fname, "工作表2");
                for (int i = 0; i < 1000; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                        continue;
                    MMeasure data = new MMeasure
                    {
                        Group = sl.GetCellValueAsString(i + 2, 1).Trim(),
                        MeasureID = sl.GetCellValueAsString(i + 2, 2).Trim(),
                        MeasureName = sl.GetCellValueAsString(i + 2, 3).Trim(),
                        Numerator = sl.GetCellValueAsString(i + 2, 4).Trim(),
                        NumeratorName = sl.GetCellValueAsString(i + 2, 5).Trim(),
                        Denominator = sl.GetCellValueAsString(i + 2, 6).Trim(),
                        DenominatorName = sl.GetCellValueAsString(i + 2, 7).Trim(),
                    };

                    GMeasures.Add(data);
                }
                sl.CloseWithoutSaving();

                sl = new SLDocument(fname, "工作表3");
                for (int i = 0; i < 10; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;
                    var multi = sl.GetCellValueAsString(i + 2, 2).Split('+').ToList();
                    if (multi.Count <= 1)
                        continue;
                    GMulgroups[sl.GetCellValueAsString(i + 2, 1)] = multi;
                }
                sl.CloseWithoutSaving();

                if (GMeasures.Count > 0)
                {
                    //MessageBox.Show("成功匯入指標數 : " + GMeasures.Count.ToString());
                    TxtBox1.Text += Environment.NewLine + "指標匯入數量 : " + GMeasures.Count + Environment.NewLine;
                }

                if (GMeasures.Count > 0)
                {
                    List<string> group = new List<string>() { "全部" };
                    foreach (var x in GMeasures)
                    {
                        if (group.Contains(x.Group))
                            continue;
                        else
                        {
                            group.Add(x.Group);
                        }
                    }
                    if (GMulgroups.Count > 0)
                    {
                        group.AddRange(GMulgroups.Select(o => o.Key));
                    }
                    Combx1.ItemsSource = group;
                    Combx2.ItemsSource = GMeasures.Select(o => o.MeasureID + ":" + o.MeasureName).ToList();
                    Combx1.SelectedIndex = 0;
                    List<string> spctype = new List<string>() { "Default", "U", "C", "P", "I", "nP", "Xbar_S", "Xbar_R" };
                    Combx3.ItemsSource = spctype;
                    Combx3.SelectedIndex = 0;
                    group.Remove("全部");
                    Combx5.ItemsSource = group;
                    Combx5.SelectedIndex = 0;
                }
                LoadDataBASE();
                GetMeasureRecords();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void LoadDataBASE()
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
                for (int i = 0; i < 1000; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;

                    if (GRecords.Count > 0 && GRecords.ContainsKey(sl.GetCellValueAsString(i + 2, 1)))
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
                    GRecords[sl.GetCellValueAsString(i + 2, 1)] = lme;

                    try
                    {
                        if (gduplicate.ContainsKey(sl.GetCellValueAsString(i + 2, 1)))
                        {
                            var glists = gduplicate.Where(o => o.Key == sl.GetCellValueAsString(i + 2, 1)).FirstOrDefault().Value;
                            foreach (var x in glists)
                            {
                                if (!GRecords.ContainsKey(x) &&
                                    GElements.Find(o => o.ElementID == x) != null)
                                {
                                    GRecords[x] = lme;
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
        public void GetMeasureRecords()
        {
            if (GMeasures.Count <= 0)
                return;
            if (GRecords.Count <= 0)
                return;
            var sortbacks = GRecords.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value.OrderByDescending(o => o.ElementRecord).ToList());

            foreach (var x in GMeasures)
            {
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
                else if (x.Denominator.Contains(" - "))
                {
                    Destatus = 2;
                    var elements = x.Denominator.Split(new[] { " - " }, StringSplitOptions.None).ToList();
                    if (elements.Count > 0)
                    {
                        foreach (var ele in elements)
                        {
                            var em = sortbacks.FirstOrDefault(o => o.Key == ele.Replace("[", "").Replace("]", "").Trim()).Value;
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
                else if (x.Numerator.Contains(" - "))
                {
                    Nustatus = 2;
                    var elements = x.Numerator.Split(new[] { " - " }, StringSplitOptions.None).ToList();
                    if (elements.Count > 0)
                    {
                        foreach (var ele in elements)
                        {
                            var em = sortbacks.FirstOrDefault(o => o.Key == ele.Replace("[", "").Replace("]", "").Trim()).Value;
                            if (em != null)
                                NumePlus.Add(em);
                        }
                    }
                }
                //將 12 + 6 個月的資料載入gmeasure
                for (int i = 0; i < 12 + 6; i++)
                {
                    double rnum = -1;
                    double rdec = -1;
                    if (x.Numerator == "1")
                        rnum = 1;
                    else if (Numes != null)
                    {
                        var nume = Numes.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-i - 1).Year
                        && o.ElementDate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                        if (nume != null)
                        {
                            if (Double.TryParse(nume.ElementRecord, out double numok))
                                rnum = numok;
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
                                rnum = deno;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                    if (x.Denominator == "1")
                        rdec = 1;
                    else if (Denos != null)
                    {
                        var deno = Denos.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-i - 1).Year
                        && o.ElementDate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                        if (deno != null)
                        {
                            if (Double.TryParse(deno.ElementRecord, out double numok))
                                rdec = numok;
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
                                rdec = deno;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                    //將分子、分母數據加回 gmeasure
                    if (rnum > -1 && rdec > -1)
                    {
                        if (!x.Records.ContainsKey(DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM")))
                        {
                            x.Records.Add(DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM"),
                            new List<double>() { rnum, rdec });
                        }
                        else
                        {
                            x.Records[DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM")] = new List<double>() { rnum, rdec };
                        }
                    }
                }
            }
        }
        public void ExportData(string type, string month)
        {
            if (!Directory.Exists(System.Environment.CurrentDirectory + @"\資料上傳"))
            {
                Directory.CreateDirectory(System.Environment.CurrentDirectory + @"\資料上傳");
            }
            if (GElements.Count <= 0)
            {
                MessageBox.Show("尚未匯入定義清單");
                return;
            }
            if (gcollect.Count <= 0)
            {
                MessageBox.Show("尚未匯入收回資料");
            }
            if (type == "TCPI")
            {
                SLDocument sl = new SLDocument();
                /*
                SLStyle style = sl.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                style.Alignment.WrapText = true;
                style.Font.FontSize = 12;
                */
                sl.SetColumnWidth(1, 10);
                sl.SetColumnWidth(2, 15);
                sl.SetColumnWidth(3, 45);
                sl.SetColumnWidth(4, 10);
                sl.SetColumnWidth(5, 15);
                sl.SetCellValue(1, 1, "日期");
                sl.SetCellValue(1, 2, "要素代碼");
                sl.SetCellValue(1, 3, "要素名稱");
                sl.SetCellValue(1, 4, "頻率");
                sl.SetCellValue(1, 5, "提報要素值");
                if (GElements.Count > 0)
                {
                    int index = 2;
                    for (int i = 0; i < GElements.Count; i++)
                    {
                        if (GElements[i].Group != type)
                            continue;
                        sl.SetCellValue(index, 1, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                        sl.SetCellValue(index, 2, GElements[i].ElementID);
                        sl.SetCellValue(index, 3, GElements[i].ElementName);
                        sl.SetCellValue(index, 4, "月");
                        if (gcollect.Count > 0)
                        {
                            var num = gcollect.FirstOrDefault(o => o.ElementID == GElements[i].ElementID && !string.IsNullOrEmpty(o.ElementRecord));
                            if (num != null)
                                sl.SetCellValue(index, 5, Convert.ToDouble(num.ElementRecord));
                        }
                        index++;
                    }
                }
                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                sl.SaveAs(Refile);
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

                SLDocument sl = new SLDocument();
                SLStyle style = sl.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                style.Font.Bold = true;
                style.Font.FontSize = 12;
                style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 176, 240), System.Drawing.Color.Black);
                SLStyle style2 = sl.CreateStyle();
                style2.Alignment.WrapText = true;
                style2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style2.Alignment.Vertical = VerticalAlignmentValues.Center;

                sl.SetColumnStyle(2, style2);
                sl.SetCellStyle(1, 1, style);
                sl.SetCellStyle(1, 2, style);
                sl.SetCellStyle(1, 3, style);
                sl.SetCellStyle(1, 4, style);
                sl.SetCellStyle(1, 5, style);
                sl.SetCellStyle(1, 6, style);
                sl.SetCellStyle(1, 7, style);
                sl.SetCellStyle(1, 8, style);
                sl.SetColumnWidth(1, 15);
                sl.SetColumnWidth(2, 45);
                sl.SetColumnWidth(3, 15);
                sl.SetColumnWidth(4, 15);
                sl.SetColumnWidth(5, 15);
                sl.SetColumnWidth(6, 15);
                sl.SetColumnWidth(7, 15);
                sl.SetColumnWidth(8, 15);
                sl.SetCellValue(1, 1, "要素代碼");
                sl.SetCellValue(1, 2, "要素名稱");
                sl.SetCellValue(1, 3, DateTime.Now.AddMonths(-1).ToString("yyyy/MM") + "(月)");
                sl.SetCellValue(1, 4, DateTime.Now.AddMonths(-2).ToString("yyyy/MM") + "(月)");
                sl.SetCellValue(1, 5, DateTime.Now.AddMonths(-3).ToString("yyyy/MM") + "(月)");
                sl.SetCellValue(1, 6, DateTime.Now.AddMonths(-4).ToString("yyyy/MM") + "(月)");
                sl.SetCellValue(1, 7, DateTime.Now.AddMonths(-5).ToString("yyyy/MM") + "(月)");
                sl.SetCellValue(1, 8, DateTime.Now.AddMonths(-6).ToString("yyyy/MM") + "(月)");

                if (GElements.Count > 0)
                {
                    //List<MElement> sortdata = new List<MElement>(GElements);
                    //sortdata.Sort((x, y) => { return x.ElementID.CompareTo(y.ElementID); });
                    int index = 2;
                    for (int i = 0; i < GElements.Count; i++)
                    {
                        if (GElements[i].Group != type)
                            continue;
                        sl.SetCellValue(index, 1, GElements[i].ElementID);
                        sl.SetCellValue(index, 2, GElements[i].ElementName);
                        if (gcollect.Count > 0)
                        {
                            var num = gcollect.FirstOrDefault(o => o.ElementID == GElements[i].ElementID && !string.IsNullOrEmpty(o.ElementRecord));
                            if (num != null)
                                sl.SetCellValue(index, 3, Convert.ToDouble(num.ElementRecord));
                        }
                        for (int j = 0; j < 5; j++)
                        {
                            var records = GRecords.FirstOrDefault(o => o.Key == GElements[i].ElementID).Value;
                            if (records == null)
                                continue;
                            var data = records.FirstOrDefault(o => o.ElementDate.Year == DateTime.Now.AddMonths(-2 - j).Year && o.ElementDate.Month == DateTime.Now.AddMonths(-2 - j).Month);
                            if (data == null)
                                continue;
                            sl.SetCellValue(index, 4 + j, Convert.ToDouble(data.ElementRecord));
                        }
                        index++;
                    }
                }

                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                sl.SaveAs(Refile);

            }
            else if (type == "THIS")
            {
                SLDocument sl = new SLDocument();
                /*
                SLStyle style = sl.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                style.Alignment.WrapText = true;
                style.Font.FontSize = 12;
                */
                sl.SetColumnWidth(1, 15);
                sl.SetColumnWidth(2, 15);
                sl.SetColumnWidth(3, 10);
                sl.SetColumnWidth(4, 10);
                sl.SetColumnWidth(5, 10);
                sl.SetColumnWidth(6, 10);
                sl.SetCellValue(1, 1, "醫院會員代碼");
                sl.SetCellValue(1, 2, "提報民國年分");
                sl.SetCellValue(1, 3, "提報月份");
                sl.SetCellValue(1, 4, "指標代碼");
                sl.SetCellValue(1, 5, "分子數據");
                sl.SetCellValue(1, 6, "分母數據");

                if (GMeasures.Count > 0)
                {
                    int index = 2;
                    Random rand = new Random();
                    for (int i = 0; i < GMeasures.Count; i++)
                    {
                        if (GMeasures[i].Group != type)
                            continue;
                        sl.SetCellValue(index, 1, "JB0005");
                        sl.SetCellValue(index, 2, DateTime.Now.AddMonths(-1).AddYears(-1911).Year);
                        sl.SetCellValue(index, 3, DateTime.Now.AddMonths(-1).Month);
                        sl.SetCellValue(index, 4, GMeasures[i].MeasureID);
                        if (gcollect.Count > 0)
                        {
                            if (GMeasures[i].Numerator == "1")
                                sl.SetCellValue(index, 5, Convert.ToDouble("1"));
                            else
                            {
                                var numer = gcollect.FirstOrDefault(o => o.ElementID == GMeasures[i].Numerator && !string.IsNullOrEmpty(o.ElementRecord));
                                if (numer != null)
                                    sl.SetCellValue(index, 5, Convert.ToDouble(numer.ElementRecord));
                            }
                            if (GMeasures[i].Denominator == "1")
                                sl.SetCellValue(index, 6, Convert.ToDouble("1"));
                            else
                            {
                                var deno = gcollect.FirstOrDefault(o => o.ElementID == GMeasures[i].Denominator && !string.IsNullOrEmpty(o.ElementRecord));
                                if (deno != null)
                                    sl.SetCellValue(index, 6, Convert.ToDouble(deno.ElementRecord));
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
                sl.SaveAs(Refile);
            }
            else if (type == "P4P")
            {
                SLDocument sl = new SLDocument();
                sl.SetColumnWidth(1, 15);
                sl.SetColumnWidth(2, 15);
                sl.SetColumnWidth(3, 45);
                sl.SetColumnWidth(4, 10);
                sl.SetColumnWidth(5, 15);
                sl.SetCellValue(1, 1, "提報月份");
                sl.SetCellValue(1, 2, "要素代碼");
                sl.SetCellValue(1, 3, "要素名稱");
                sl.SetCellValue(1, 4, "頻率");
                sl.SetCellValue(1, 5, "提報要素值");

                SLStyle style = sl.CreateStyle();
                style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Cyan, System.Drawing.Color.Black);
                sl.SetCellStyle(1, 1, style);
                sl.SetCellStyle(1, 2, style);
                sl.SetCellStyle(1, 3, style);
                sl.SetCellStyle(1, 4, style);
                sl.SetCellStyle(1, 5, style);

                if (GElements.Count > 0)
                {
                    int index = 2;
                    GElements.ForEach((x) =>
                    {
                        if (x.Group == type)
                        {
                            sl.SetCellValue(index, 1, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                            sl.SetCellValue(index, 2, x.ElementID);
                            sl.SetCellValue(index, 3, x.ElementName);
                            sl.SetCellValue(index, 4, "月");
                            if (gcollect.Count > 0)
                            {
                                var number = gcollect.FirstOrDefault(o => o.ElementID == x.ElementID && !string.IsNullOrEmpty(o.ElementRecord));
                                if (number != null)
                                    sl.SetCellValue(index, 5, Convert.ToDouble(number.ElementRecord));
                            }
                            index++;
                        }
                    });
                    /*
                    for (int i = 0; i < GElements.Count; i++)
                    {
                        if (GElements[i].Group != type)
                            continue;
                        sl.SetCellValue(index, 1, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                        sl.SetCellValue(index, 2, GElements[i].ElementID);
                        sl.SetCellValue(index, 3, GElements[i].ElementName);
                        sl.SetCellValue(index, 4, "月");
                        if (gcollect.Count > 0)
                        {
                            var number = gcollect.FirstOrDefault(o => o.ElementID == GElements[i].ElementID && !string.IsNullOrEmpty(o.ElementRecord));
                            if (number != null)
                                sl.SetCellValue(index, 5, Convert.ToDouble(number.ElementRecord));
                        }
                        index++;
                    }
                    */
                }
                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                sl.SaveAs(Refile);
            }
            if (GMulgroups.Count > 0)
            {
                foreach (var x in GMulgroups)
                {
                    if (x.Key == "內科")
                    {
                        List<MElement> Geles = new List<MElement>();
                        var fname = @"\指標數據總資料" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + x.Key + "(群組資料).xlsx";
                        var Gmeas = GMeasures.Where(o => o.Group == "群組指標").ToList();
                        foreach (var ele in Gmeas)
                        {
                            MMeasure tp = new MMeasure();
                            List<MMeasure> tps = new List<MMeasure>();
                            foreach (var station in x.Value)
                            {
                                if (station.Contains(ele.MeasureID))
                                {
                                    tps.Add(ele);
                                }
                            }
                        }
                        using (SLDocument sl = new SLDocument())
                        {

                            //sl.SaveAs(fpath + fname);
                        }
                    }
                }
            }
            MessageBox.Show("轉檔結束");
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
                sdata = GElements.Where(x => x.Group == gtype).ToList();
            else
                sdata = GElements;
        }
        #endregion

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
            List<string> uncollect = new List<string>();
            List<string> collect = new List<string>();

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
                    for (int i = 2; i < 1000; i++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 1)))
                            break;
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 3)))
                            continue;
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 5)))
                        {
                            if (!uncollect.Contains(sl.GetCellValueAsString(i, 3).Trim()) && !collect.Contains(sl.GetCellValueAsString(i, 3).Trim()))
                                uncollect.Add(sl.GetCellValueAsString(i, 3).Trim());
                            continue;
                        }
                        // 匯入資料若無法轉換成double，表示資料有錯誤
                        if (!Double.TryParse(sl.GetCellValueAsString(i, 5).Trim(), out double rd))
                        {
                            MessageBox.Show("指標: " + sl.GetCellValueAsString(i, 3).Trim() + "之數值有誤 (" + sl.GetCellValueAsString(i, 5).Trim() + ")");
                            if (!uncollect.Contains(sl.GetCellValueAsString(i, 3).Trim()) && !collect.Contains(sl.GetCellValueAsString(i, 3).Trim()))
                                uncollect.Add(sl.GetCellValueAsString(i, 3).Trim());
                            continue;
                        }
                        if (!collect.Contains(sl.GetCellValueAsString(i, 3).Trim()))
                            collect.Add(sl.GetCellValueAsString(i, 3).Trim());
                        if (uncollect.Contains(sl.GetCellValueAsString(i, 3).Trim()))
                            uncollect.Remove(sl.GetCellValueAsString(i, 3).Trim());
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

                if (uncollect.Count > 0)
                {
                    TxtBox1.Text += Environment.NewLine + string.Format("未成功收回指標 ({0}) 個 : {1}", uncollect.Count, string.Join(",", uncollect)) +
                        Environment.NewLine;
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
                                var glist = gcollect.FirstOrDefault(o => o.ElementID == x.Key);
                                if (glist == null)
                                {
                                    List<string> elements = new List<string>();
                                    var complexs = x.Value.FirstOrDefault(o => o.Contains("+"));
                                    if (complexs != null)
                                        elements = complexs.Split('+').ToList();
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
                                            if (double.TryParse(em.ElementRecord, out double record))
                                                total += record;
                                            else
                                            {
                                                MessageBox.Show("error" + ele.Trim());
                                                collected = false;
                                                break;
                                            }
                                            collected = true;
                                        }
                                    }
                                    else
                                    {
                                        x.Value.ForEach((o) =>
                                        {
                                            var em = gcollect.FirstOrDefault(z => z.ElementID == o);
                                            if (!collected && em != null && double.TryParse(em.ElementRecord, out double record))
                                            {
                                                total = record;
                                                collected = true;
                                            }
                                        });
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
                                        if (!o.Contains("+") && gcollect.FirstOrDefault(z => z.ElementID == o) == null)
                                        {
                                            gcollect.Add(new MElement
                                            {
                                                ElementID = o,
                                                ElementRecord = glist.ElementRecord,
                                                ElementDate = DateTime.Now.AddMonths(-1)
                                            });
                                        }
                                    });
                                }
                            }
                        }
                        /// 再次檢查&合併相同數值之要素
                        foreach (var x in gduplicate)
                        {
                            var glist = gcollect.FirstOrDefault(o => o.ElementID == x.Key);
                            if (glist == null)
                                continue;
                            x.Value.ForEach((o) =>
                            {
                                if (!o.Contains("+") && gcollect.FirstOrDefault(z => z.ElementID == o) == null)
                                {
                                    gcollect.Add(new MElement
                                    {
                                        ElementID = o,
                                        ElementRecord = glist.ElementRecord,
                                        ElementDate = DateTime.Now.AddMonths(-1)
                                    });
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
                    TxtBox1.Text += Environment.NewLine + string.Format("指標收回數量 : {0}/{1} ({2}%)", gcollect.Count, GElements.Count, gcollect.Count * 100 / GElements.Count) +
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
                            if (dts.Month != DateTime.Now.AddMonths(-1).Month ||
                                dts.Year != DateTime.Now.AddMonths(-1).Year)
                            {
                                //sl.CopyColumn(2, 100, 3);
                                sl.InsertColumn(2, 1);
                                sl.SetCellValue(1, 2, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));
                            }

                            foreach (var x in gcollect)
                            {
                                SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();
                                int slrows = wsstats.EndRowIndex;
                                bool oldele = false;
                                for (int i = 0; i < slrows + 10; i++)
                                {
                                    //if (string.IsNullOrWhiteSpace(sl.GetCellValueAsString(i + 2, 1)))
                                    //    break;
                                    if (sl.GetCellValueAsString(i + 2, 1) == x.ElementID)
                                    {
                                        sl.SetCellValue(i + 2, 2, Convert.ToDouble(x.ElementRecord));
                                        oldele = true;
                                        break;
                                    }
                                }
                                if (!oldele)
                                {
                                    sl.SetCellValue(slrows + 1, 1, x.ElementID);
                                    sl.SetCellValue(slrows + 1, 2, Convert.ToDouble(x.ElementRecord));
                                }
                            }
                            sl.SaveAs(fpath + @"\指標收集存檔總檔" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx");
                            //sl.SaveAs(fpath + fname2);
                        }
                    }
                    // 將收集之資料匯回資料庫中
                    foreach (var x in gcollect)
                    {
                        if (GRecords.ContainsKey(x.ElementID))
                        {
                            var dataex = GRecords[x.ElementID].FirstOrDefault(o => o.ElementID == x.ElementID && o.ElementDate == x.ElementDate);
                            if (dataex != null)
                                GRecords[x.ElementID].Remove(dataex);
                            GRecords[x.ElementID].Add(x);
                        }
                        else
                        {
                            GRecords[x.ElementID] = new List<MElement>();
                            GRecords[x.ElementID].Add(x);
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
            if (GElements.Count <= 0)
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
            List<MElement> newda = new List<MElement> (GElements);
            if (GMulgroups.Count <= 0)
                newda.Sort((x, y) => { return x.ElementName.CompareTo(y.ElementName); });
            else
                newda.Sort((x, y) => { return x.ElementID.CompareTo(y.ElementID); });
            var newdata = newda.GroupBy(o => o.Depart)
                    .ToDictionary(o => o.Key, o => o.ToList());

            var unitcounts = GElements.Where(o => !SameEle.Contains(o.ElementID))
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

                        if (GRecords.Count > 0 && GRecords.ContainsKey(x.Value[i].ElementID))
                        {
                            for (int j = 2; j < 7; j++)
                            {
                                var data = GRecords[x.Value[i].ElementID].Find(o => o.ElementDate.Month == DateTime.Now.AddMonths(-j).Month
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
            if (GMeasures.Count <= 0)
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
                        for (int i = 0; i < 1000; i++)
                        {
                            if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                                break;
                            if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                                continue;
                            if (!DateTime.TryParse(((sl.GetCellValueAsInt32(i + 2, 2) + 1911).ToString() + "/" + sl.GetCellValueAsString(i + 2, 3)), out DateTime dts))
                                continue;
                            var measureid = GMeasures.Where(o => o.MeasureID == sl.GetCellValueAsString(i + 2, 4)).FirstOrDefault();
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
                                    if (!GRecords.ContainsKey(measureid.Numerator))
                                    {

                                        lmeNume.Add(dataNume);
                                        GRecords[measureid.Numerator] = lmeNume;
                                    }
                                    else
                                    {
                                        var meNume = GRecords[measureid.Numerator].FirstOrDefault(o => o.ElementDate == dts);
                                        if (meNume == null)
                                        {
                                            GRecords[measureid.Numerator].Add(dataNume);
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
                                    if (!GRecords.ContainsKey(measureid.Denominator))
                                    {

                                        lmeDeno.Add(dataDeno);
                                        GRecords[measureid.Denominator] = lmeDeno;
                                    }
                                    else
                                    {
                                        var meDeno = GRecords[measureid.Denominator].FirstOrDefault(o => o.ElementDate == dts);
                                        if (meDeno == null)
                                        {
                                            GRecords[measureid.Denominator].Add(dataDeno);
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
            if (GRecords.Count <= 0)
                return;
            /*
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
            */

            var sortbacks = GRecords.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);

            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標要素總存檔" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx";
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
                    GElements.ForEach((x) =>
                    {
                        sl.SetCellValue(index + 2, 1, x.ElementID);
                        index++;
                    }) ;
                    SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();
                    int slrows = wsstats.EndRowIndex;
                    for (int i = 0; i < slrows; i++)
                    {
                        var x = sortbacks.FirstOrDefault(o => o.Key == sl.GetCellValueAsString(i + 2, 1)).Value;
                        if (x != null)
                        {
                            for (int j = 0; j < 12 + 6; j++)
                            {
                                if (!DateTime.TryParse(sl.GetCellValueAsString(1, j + 2), out DateTime dts))
                                    continue;
                                var y = x.FirstOrDefault(o => o.ElementDate.Year == dts.Year && o.ElementDate.Month == dts.Month);
                                if (y != null && Double.TryParse(y.ElementRecord, out double num))
                                {
                                    sl.SetCellValue(i + 2, j + 2, num);
                                }
                            }
                        }
                    }
                    /*
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
                    */
                    
                    sl.SaveAs(fpath + fname);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageBox.Show("指標要素匯出結束");
        }

        private void BT_EXPORT_MEASURE(object sender, RoutedEventArgs e)
        {
            if (GMeasures.Count <= 0)
                return;
            if (GRecords.Count <= 0)
                return;
            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標數據總資料" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ".xlsx";

            var sortbacks = GRecords.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value.OrderByDescending(o => o.ElementRecord).ToList());
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
                    foreach (var x in GMeasures)
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
                        else if (x.Denominator.Contains(" - "))
                        {
                            Destatus = 2;
                            var elements = x.Denominator.Split(new[] { " - " }, StringSplitOptions.None).ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele.Replace("[", "").Replace("]", "").Trim()).Value;
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
                        else if (x.Numerator.Contains(" - "))
                        {
                            Nustatus = 2;
                            var elements = x.Numerator.Split(new[] { " - " }, StringSplitOptions.None).ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele.Replace("[", "").Replace("]", "").Trim()).Value;
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
                            {
                                x.Records.Add(DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM"),
                                new List<double>() { sl.GetCellValueAsDouble(index + 1, i + 4), sl.GetCellValueAsDouble(index + 2, i + 4) });
                            }
                            else
                            {
                                x.Records[DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM")] = new List<double>() { sl.GetCellValueAsDouble(index + 1, i + 4), sl.GetCellValueAsDouble(index + 2, i + 4) };
                            }
                            

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
            MessageBox.Show("指標數據匯出結束");
        }

        private void BT_EXPORT_CHART(object sender, RoutedEventArgs e)
        {
            if (GMeasures.Count < 0)
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
            var smeasure = GMeasures.FirstOrDefault(o => o.MeasureID == selmeasure);
            if (smeasure == null || smeasure.Records.Count <= 0)
                return;
            if (!DateTime.TryParse(Combx4_Month_Ed.SelectedValue.ToString(), out DateTime ed)
                || !DateTime.TryParse(Combx4_Month_ST.SelectedValue.ToString(), out DateTime st)
                || ed <= st.AddMonths(3))
                return;
            int cMonths = ed.Year * 12 + ed.Month - (st.Year * 12 + st.Month) + 1;
            if (cMonths < 3)
            {
                MessageBox.Show("選取月份太短");
                return;
            }

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
                    sl.SetCellValue(2, 2, Combx2.SelectedValue.ToString() + (spctype == SPCtype.P ? "(‰)" : ""));
                    sl.SetCellValue(3, 2, "平均值");
                    sl.SetCellValue(4, 2, "管制圖上限 (2α)");
                    sl.SetCellValue(5, 2, "管制圖下限 (2α)");
                    sl.SetCellValue(6, 2, "管制圖上限 (3α)");
                    sl.SetCellValue(7, 2, "管制圖下限 (3α)");
                    sl.SetCellValue(8, 2, "分子" + (char)10 + smeasure.NumeratorName);
                    sl.SetCellValue(9, 2, "分母" + (char)10 + smeasure.DenominatorName);
                    //取得數值資料
                    for (int i = 0; i < cMonths; i++)
                    {
                        var srecord = smeasure.Records.Where(o => DateTime.TryParse(o.Key, out DateTime sda)
                        && sda.Year == DateTime.Now.AddMonths(-cMonths - 1 + i).Year
                        && sda.Month == DateTime.Now.AddMonths(-cMonths - 1 + i).Month).ToDictionary(o => o.Key, o => o.Value);
                        if (srecord == null || srecord.Count <= 0)
                            continue;

                        sl.SetCellValue(1, 3 + i, srecord.Keys.FirstOrDefault());
                        sl.SetCellValue(8, 3 + i, srecord.FirstOrDefault().Value[0]);
                        sl.SetCellValue(9, 3 + i, srecord.FirstOrDefault().Value[1]);
                        //sl.SetCellValue(1, 3 + i, sl.GetCellValueAsString(1, 3 + 12 - i));
                        //sl.SetCellValue(8, 3 + i, sl.GetCellValueAsDouble(3, 3 + 12 - i));
                        //sl.SetCellValue(9, 3 + i, sl.GetCellValueAsDouble(4, 3 + 12 - i));
                    }
                    /*
                     * 數值版
                     */
                    double avg = 0;
                    double sig = 0;
                    double mea = 0;
                    double amp = 1;
                    double u_p = 0;
                    double n_n = 0;

                    double sum_a = 0;
                    double sum_b = 0;
                    for (int i = 0; i < cMonths; i++)
                    {
                        sum_a += sl.GetCellValueAsDouble(8, 3 + i);
                        sum_b += sl.GetCellValueAsDouble(9, 3 + i);
                    }
                    if (spctype == SPCtype.C)
                    {
                        u_p = sum_a / cMonths;
                        n_n = cMonths;
                        //sl.SetCellValue(2, 1, "c");
                    }
                    else
                    {
                        u_p = sum_a / sum_b;
                        n_n = sum_b / cMonths;
                        //sl.SetCellValue(2, 1, "u or p");
                    }

                    //sl.SetCellValue(3, 1, u_p);
                    //sl.SetCellValue(4, 1, "n");
                    //sl.SetCellValue(5, 1, n_n);
                    amp = spctype == SPCtype.P ? 1000 : 1;
                    avg = spctype == SPCtype.nP ? u_p * n_n : u_p;

                    for (int i = 0; i < cMonths; i++)
                    {
                        int index = 3 + i;
                        if (spctype == SPCtype.P)
                        {
                            sig = Math.Sqrt(u_p * (1 - u_p) / sl.GetCellValueAsDouble(9, index));
                            mea = (sl.GetCellValueAsDouble(8, index) / sl.GetCellValueAsDouble(9, index)) * amp;
                        }
                        else if (spctype == SPCtype.U)
                        {
                            sig = Math.Sqrt(u_p / sl.GetCellValueAsDouble(9, index));
                            mea = (sl.GetCellValueAsDouble(8, index) / sl.GetCellValueAsDouble(9, index)) * amp;
                        }
                        else if (spctype == SPCtype.C)
                        {
                            sig = Math.Sqrt(u_p);
                            mea = sl.GetCellValueAsDouble(8, index);
                        }
                        else if (spctype == SPCtype.nP)
                        {
                            sig = Math.Sqrt(n_n * u_p * (1 - u_p));
                            mea = sl.GetCellValueAsDouble(8, index);
                        }
                        else //暫時借用 U 圖
                        {
                            sig = Math.Sqrt(u_p / sl.GetCellValueAsDouble(9, index));
                            mea = (sl.GetCellValueAsDouble(8, index) / sl.GetCellValueAsDouble(9, index)) * amp;
                        }
                        if (sl.GetCellValueAsDouble(9, index) == 0)
                            sl.SetCellValue(2, index, 0);
                        else
                            sl.SetCellValue(2, index, mea);
                        sl.SetCellValue(3, index, avg * amp);
                        sl.SetCellValue(4, index, (avg + (sig * 2)) * amp);
                        sl.SetCellValue(5, index, (avg - (sig * 2)) * amp);
                        sl.SetCellValue(6, index, (avg + (sig * 3)) * amp);
                        sl.SetCellValue(7, index, (avg - (sig * 3)) * amp);
                    }
                    /*
                     * 公式版
                     */
                     /*
                    if (spctype == SPCtype.C)
                    {
                        sl.SetCellValue(2, 1, "c");
                        sl.SetCellValue(3, 1, string.Format("=SUM(C8:{0}8) / {1}", alphabet[cMonths + 1], cMonths));
                        sl.SetCellValue(4, 1, "n");
                        sl.SetCellValue(5, 1, cMonths);
                    }
                    else
                    {
                        sl.SetCellValue(2, 1, "u or p");
                        sl.SetCellValue(3, 1, string.Format("=SUM(C8:{0}8) / SUM(C9:{0}9)", alphabet[cMonths + 1]));
                        sl.SetCellValue(4, 1, "n");
                        sl.SetCellValue(5, 1, string.Format("=AVERAGE(C9:{0}9)", alphabet[cMonths + 1]));
                    }

                    for (int i = 0; i < cMonths; i++)
                    {
                        string sigma = string.Empty;
                        string amplify = string.Empty;
                        string average = "A3";
                        string measure = string.Empty;
                        if (spctype == SPCtype.P)
                        {
                            amplify = " * 1000";
                            sigma = string.Format("SQRT(A3 * (1 - A3) / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        else if (spctype == SPCtype.U)
                        {
                            amplify = string.Empty;
                            sigma = string.Format("SQRT(A3 / {0}9)", alphabet[2 + i]);
                            measure = string.Format("=({0}8 / {0}9){1}", alphabet[2 + i], amplify);
                        }
                        else if (spctype == SPCtype.C)
                        {
                            amplify = string.Empty;
                            sigma = "SQRT(A3)";
                            measure = string.Format("=({0}8)", alphabet[2 + i]);
                        }
                        else if (spctype == SPCtype.nP)
                        {
                            average = "A5 * A3";
                            amplify = string.Empty;
                            sigma = "SQRT(A5 * A3 * (1 - A3))";
                            measure = string.Format("=({0}8)", alphabet[2 + i]);
                        }
                        else //暫時借用 U 圖
                        {
                            amplify = string.Empty;
                            sigma = string.Format("SQRT(A3 / {0}9)", alphabet[2 + i]);
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
                    */

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
                    double fChartWidth = Math.Max(13, cMonths + 1);
                    SLChart chart;
                    chart = sl.CreateChart("B1", alphabet[cMonths + 1] + "7");
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

        private void Combx1_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (GMeasures.Count <= 0)
                return;

            if (Combx1.SelectedIndex == 0)
            {
                Combx2.ItemsSource = GMeasures.Select(o => o.MeasureID + ":" + o.MeasureName).ToList();
            }
            else
            {
                Combx2.ItemsSource = GMeasures.Where(o => o.Group == Combx1.SelectedValue.ToString()).Select(o => o.MeasureID + ":" + o.MeasureName).ToList();
            }
        }

        private void EXPORT_UPLOAD(object sender, RoutedEventArgs e)
        {
            try
            {
                ExportData(Combx5.SelectedValue.ToString(), Combx4_Month.SelectedValue.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BT_CONVERT_2_XLSX(object sender, RoutedEventArgs e)
        {
            string folderName = System.Environment.CurrentDirectory + @"\資料匯總";
            try
            {
                foreach (var finame in System.IO.Directory.GetFileSystemEntries(folderName))
                {
                    if (System.IO.Path.GetExtension(finame) != ".xls")
                        continue;
                    Spire.Xls.Workbook wk = new Spire.Xls.Workbook();
                    wk.LoadFromFile(finame);
                    wk.SaveToFile(folderName + @"\" + System.IO.Path.GetFileNameWithoutExtension(finame) + ".xlsx", Spire.Xls.ExcelVersion.Version2010);

                    string targetpath = folderName + @"\" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + @"\";
                    string targetname = targetpath + System.IO.Path.GetFileName(finame);
                    if (!Directory.Exists(targetpath))
                    {
                        Directory.CreateDirectory(targetpath);
                    }
                    if (System.IO.File.Exists(targetname))
                    {
                        System.IO.File.Delete(targetname);
                    }
                    System.IO.File.Move(finame, targetname);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BT_CONVERT_2_XLS(object sender, RoutedEventArgs e)
        {
            string folderName = System.Environment.CurrentDirectory + @"\資料收集";
            try
            {
                foreach (var finame in System.IO.Directory.GetFileSystemEntries(folderName))
                {
                    if (System.IO.Path.GetExtension(finame) != ".xlsx")
                        continue;
                    Spire.Xls.Workbook wk = new Spire.Xls.Workbook();
                    wk.LoadFromFile(finame);
                    wk.SaveToFile(folderName + @"\xls\" + System.IO.Path.GetFileNameWithoutExtension(finame) + ".xls", Spire.Xls.ExcelVersion.Version97to2003);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BT_OPEN_CHART(object sender, RoutedEventArgs e)
        {
            Chart ct = new Chart();
            ct.GElements = GElements;
            ct.GMeasures = GMeasures;
            this.Hide();
            try
            {
                ct.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            this.Show();
        }
    }
}
