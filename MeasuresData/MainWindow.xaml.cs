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
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        public List<DiliveryData> gdata = new List<DiliveryData>();
        public List<Measurement> gmeasure = new List<Measurement>();
        public List<MElements> gcollect = new List<MElements>();
        public Dictionary<string, List<MElements>> gbackups = new Dictionary<string, List<MElements>>();
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
                    DiliveryData data = new DiliveryData
                    {
                        Group = sl.GetCellValueAsString(i + 2, 1).Trim(),
                        Depart = sl.GetCellValueAsString(i + 2, 2).Trim(),
                        MeasureID = sl.GetCellValueAsString(i + 2, 3).Trim(),
                        MeasureName = sl.GetCellValueAsString(i + 2, 5).Trim(),
                        User = sl.GetCellValueAsString(i + 2, 7).Trim()
                    };
                    for (int j = 12; j < 15; j++)
                    {
                        string content = sl.GetCellValueAsString(i + 2, j).Trim();
                        if (string.IsNullOrEmpty(content))
                            break;

                        if (SameEle.Contains(content) || (gduplicate.ContainsKey(content) && gduplicate[content].Contains(data.MeasureID)))
                            continue;

                        SameEle.Add(content);

                        if (gduplicate.ContainsKey(data.MeasureID))
                        {
                            if (!gduplicate[data.MeasureID].Contains(content))
                            {
                                gduplicate[data.MeasureID].Add(content);
                            }
                        }
                        else
                        {
                            gduplicate.Add(data.MeasureID, new List<string>() { content });
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
                    Measurement data = new Measurement
                    {
                        Group = sl2.GetCellValueAsString(i + 2, 1).Trim(),
                        MeasureID = sl2.GetCellValueAsString(i + 2, 2).Trim(),
                        MeasureName = sl2.GetCellValueAsString(i + 2, 3).Trim(),
                        Numerator = sl2.GetCellValueAsString(i + 2, 4).Trim(),
                        Denominator = sl2.GetCellValueAsString(i + 2, 6).Trim()
                    };

                    gmeasure.Add(data);
                }

                sl2.CloseWithoutSaving();

                MessageBox.Show("匯入成功 : " + gmeasure.Count.ToString());

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
                    //if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                    //    continue;

                    if (gbackups.Count > 0 && gbackups.ContainsKey(sl.GetCellValueAsString(i + 2, 1)))
                    {
                        continue;
                    }
                    List<MElements> lme = new List<MElements>();
                    List<string> duplicate = new List<string>();
                    for (int j = 0; j < 12; j++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(1, j + 2)))
                            break;
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(2, j + 2)))
                            continue;
                        if (!DateTime.TryParse(sl.GetCellValueAsDateTime(1, j + 2).ToShortDateString(), out DateTime dts))
                            continue;
                        if (dts > DateTime.Now.AddMonths(-1 - 12) && dts < DateTime.Now)
                        {
                            if (duplicate.Contains(dts.ToString("yyyy/MM")))
                                continue;
                            else
                                duplicate.Add(dts.ToString("yyyy/MM"));

                            MElements data = new MElements
                            {
                                Element = sl.GetCellValueAsString(i + 2, 1).Trim(),
                                ElementData = sl.GetCellValueAsString(i + 2, j + 2).Trim(),
                                Eledate = dts
                            };
                            lme.Add(data);
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
                                    gdata.Find(o => o.MeasureID == x) != null)
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
                sl.Dispose();
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
                        nsl.SetCellValue(index, 1, DateTime.Now.ToString("yyyy/M"));
                        nsl.SetCellValue(index, 2, gdata[i].MeasureID);
                        nsl.SetCellValue(index, 3, gdata[i].MeasureName);
                        nsl.SetCellValue(index, 4, "月");
                        if (gcollect.Count > 0)
                        {
                            var number = from num in gcollect
                                         where num.Element == gdata[i].MeasureID && !string.IsNullOrEmpty(num.ElementData)
                                         select num;
                            if (number != null && number.ToList().Count > 0)
                                nsl.SetCellValue(index, 5, number.ToList().First().ElementData);
                        }
                        index++;
                    }
                }
                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                nsl.SaveAs(Refile);
                nsl.Dispose();
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
                nsl.Dispose();
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
                    sortdata.Sort((x, y) => { return x.MeasureID.CompareTo(y.MeasureID); });
                    int index = 2;
                    for (int i = 0; i < sortdata.Count; i++)
                    {
                        if (sortdata[i].Group != type)
                            continue;
                        nsl.SetCellValue(index, 1, sortdata[i].MeasureID);
                        nsl.SetCellValue(index, 2, sortdata[i].MeasureName);
                        if (gcollect.Count > 0)
                        {
                            var number = from num in gcollect
                                         where num.Element == gdata[i].MeasureID && !string.IsNullOrEmpty(num.ElementData)
                                         select num;
                            if (number != null && number.ToList().Count > 0)
                                nsl.SetCellValue(index, 3, number.ToList().First().ElementData);
                        }
                        index++;
                    }
                }

                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                nsl.SaveAs(Refile);
                nsl.Dispose();

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
                        nsl.SetCellValue(index, 2, DateTime.Now.AddYears(-1911).Year);
                        nsl.SetCellValue(index, 3, DateTime.Now.Month);
                        nsl.SetCellValue(index, 4, gmeasure[i].MeasureID);
                        if (gcollect.Count > 0)
                        {
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

                        }
                        //nsl.SetCellValue(index, 5, gmeasure[i].Numerator);
                        //nsl.SetCellValue(index, 6, gmeasure[i].Denominator);
                        index++;
                    }
                }
                string Refile = System.Environment.CurrentDirectory + @"\資料上傳\" + type + " (" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM") + ")" + ".xlsx";
                nsl.SaveAs(Refile);
                nsl.Dispose();
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

                var list = cn.Query<MElements>(
                    "SELECT * FROM MeasureTable WHERE MeasureID=@catg", new { catg = "EDP005-01" });

                foreach (var item in list)
                {
                    TxtBox1.Text += item.Eledate.Year + "/" + item.Eledate.Month + "=" + item.ElementData + Environment.NewLine;
                }

            }

        }
        public class DiliveryData
        {
            public string Group { get; set; }
            public string Depart { get; set; }
            public string MeasureID { get; set; }
            public string MeasureName { get; set; }
            public string Frequency { get; set; }
            public string User { get; set; }

            public List<string> SameID = new List<string>();
        }

        public class Measurement
        {
            public string Group { get; set; }
            public string MeasureID { get; set; }
            public string MeasureName { get; set; }
            public string Numerator { get; set; }
            public string Denominator { get; set; }
            public string Threshold { get; set; }
            public string Frequency { get; set; }
            public string User { get; set; }
        }

        public class MElements
        {
            public string Element { get; set; }
            public DateTime Eledate { get; set; }
            public string ElementData { get; set; }
            public Dictionary<string, string> PreDdatas { get; set; }
            public MElements()
            {
                PreDdatas = new Dictionary<string, string>();
                Eledate = new DateTime();
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

        private void BT_IMPORT_RESULT(object sender, RoutedEventArgs e)
        {
            /*
            if (!Directory.Exists(System.Environment.CurrentDirectory + @"\要素備份"))
            {
                Directory.CreateDirectory(System.Environment.CurrentDirectory + @"\要素備份");
            }
            string fname = System.Environment.CurrentDirectory + @"\要素備份\指標收集存檔.xlsx";

            SLDocument sl = new SLDocument(!System.IO.File.Exists(fname) ? "" : fname);
            if (!string.IsNullOrEmpty(sl.GetCellValueAsString(1, 5)))
            {

            }
            */
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
                    if (!DateTime.TryParse(sl.GetCellValueAsDateTime(1, 5).ToShortDateString(), out DateTime time))
                        continue;
                    if (time.Year != DateTime.Now.AddMonths(-1).Year
                        || time.Month != DateTime.Now.AddMonths(-1).Month)
                        continue;
                    for (int i = 2; i < 500; i++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 1)))
                            break;
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i, 3)))
                            continue;
                        MElements data = new MElements
                        {
                            Element = sl.GetCellValueAsString(i, 3).Trim(),
                            ElementData = sl.GetCellValueAsString(i, 5).Trim(),
                            Eledate = time
                        };
                        if (gcollect.Count > 0)
                        {
                            bool dup = false;
                            foreach (var x in gcollect)
                            {
                                if (data.Element == x.Element)
                                {
                                    duplicate.Add(x.Element);
                                    dup = true;
                                    break;
                                }
                            }
                            if (!dup)
                                gcollect.Add(data);
                        }
                        else
                            gcollect.Add(data);

                        try
                        {
                            if (gduplicate.ContainsKey(data.Element))
                            {
                                var glists = gduplicate.Where(o => o.Key == data.Element).FirstOrDefault().Value;
                                foreach (var x in glists)
                                {
                                    if (gcollect.Find(o => o.Element == x.ToString()) == null &&
                                        gdata.Find(o => o.MeasureID == x.ToString()) != null)
                                    {
                                        //MessageBox.Show(x.ToString());
                                        MElements dupdata = new MElements
                                        {
                                            Element = x.ToString(),
                                            ElementData = data.ElementData,
                                            Eledate = data.Eledate
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
                    }

                    sl.CloseWithoutSaving();
                    sl.Dispose();
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
            if (gcollect.Count > 0)
            {
                TxtBox1.Text += Environment.NewLine + string.Format("指標收回數量 : {0}/{1} ({2}%)", gcollect.Count, gdata.Count, gcollect.Count * 100 / gdata.Count) +
                    Environment.NewLine;

                gcollect.Sort((x, y) => { return x.Element.CompareTo(y.Element); });

                string fpath = Environment.CurrentDirectory + @"\要素備份";
                if (!Directory.Exists(fpath))
                {
                    Directory.CreateDirectory(fpath);
                }
                string fname = @"\指標收集存檔" + DateTime.Now.AddMonths(-1).ToString("yyyy-M") + ".xlsx";
                string fname2 = @"\指標收集存檔總檔.xlsx";
                using (SLDocument nsl = new SLDocument())
                {

                    nsl.SetColumnWidth(1, 15);
                    nsl.SetColumnWidth(2, 15);
                    nsl.SetCellValue(1, 1, "指標要素");
                    nsl.SetCellValue(1, 2, DateTime.Now.AddMonths(-1).ToString("yyyy/M"));
                    for (int i = 0; i < gcollect.Count; i++)
                    {
                        nsl.SetCellValue(i + 2, 1, gcollect[i].Element);
                        nsl.SetCellValue(i + 2, 2, gcollect[i].ElementData);
                    }

                    nsl.SaveAs(fpath + fname);

                    if (!System.IO.File.Exists(fpath + fname2))
                    {
                        SLDocument sl2 = new SLDocument(fpath + fname);
                        sl2.SaveAs(fpath + fname2);
                        sl2.Dispose();
                        return;
                    }
                    nsl.Dispose();
                }

                if (System.IO.File.Exists(fpath + fname2))
                {
                    using (SLDocument sl = new SLDocument(fpath + fname2))
                    {
                        if (!DateTime.TryParse(sl.GetCellValueAsDateTime(1, 2).ToShortDateString(), out DateTime dts))
                        {
                            sl.CloseWithoutSaving();
                            return;
                        }
                        MessageBox.Show(dts.ToShortDateString());
                        if ((dts.Month == DateTime.Now.AddMonths(-1).Month &&
                            dts.Year == DateTime.Now.AddMonths(-1).Year) || dts >= DateTime.Now)
                            return;
                        sl.CopyColumn(2, 100, 3);
                        sl.SetCellValue(1, 2, DateTime.Now.AddMonths(-1).ToString("yyyy/MM"));

                        for (int i = 0; i < gcollect.Count; i++)
                        {
                            for (int j = 0; j < 500; j++)
                            {
                                if (string.IsNullOrWhiteSpace(sl.GetCellValueAsString(j + 2, 1)))
                                    break;
                                if (sl.GetCellValueAsString(j + 2, 1) == gcollect[i].Element)
                                {
                                    sl.SetCellValue(j + 2, 2, gcollect[i].ElementData);
                                    break;
                                }
                            }
                        }
                        sl.SaveAs(fpath + fname2);
                        sl.Dispose();
                    }
                }
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
            newda.Sort((x, y) => { return x.MeasureName.CompareTo(y.MeasureName); });
            var newdata = newda.GroupBy(o => o.Depart)
                    .ToDictionary(o => o.Key, o => o.ToList());

            var unitcounts = gdata.Where(o => !SameEle.Contains(o.MeasureID))
                .GroupBy(o => o.Depart)
                .ToDictionary(o => o.Key, o => o.ToList().Count);
            TxtBox1.Text += Environment.NewLine + "指標收集單位數 : " + unitcounts.Count +
                Environment.NewLine + string.Join(",", unitcounts) + Environment.NewLine;

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
                    if (SameEle.Count > 0 && SameEle.Contains(x.Value[i].MeasureID))
                    {
                        continue;
                    }
                    nsl.SetCellValue(index, 1, x.Value[i].Group);
                    nsl.SetCellValue(index, 2, x.Value[i].Depart);
                    nsl.SetCellValue(index, 3, x.Value[i].MeasureID);
                    nsl.SetCellValue(index, 4, x.Value[i].MeasureName);

                    nsl.SetCellStyle(index, 5, style);

                    if (gbackups.Count > 0 && gbackups.ContainsKey(x.Value[i].MeasureID))
                    {
                        for (int j = 2; j < 7; j++)
                        {
                            var data = gbackups[x.Value[i].MeasureID].Find(o => o.Eledate.Month == DateTime.Now.AddMonths(-j).Month
                            && o.Eledate.Year == DateTime.Now.AddMonths(-j).Year);
                            if (data == null)
                                continue;

                            nsl.SetCellValueNumeric(index, j + 4, data.ElementData);
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
                nsl.Dispose();
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
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = Environment.CurrentDirectory;
            dlg.Title = "選取資料檔";
            dlg.Filter = "xlsx files (*.*)|*.xlsx";
            if (dlg.ShowDialog() != true)
                return;
            if (!System.IO.File.Exists(dlg.FileName))
                return;
            using (SLDocument sl = new SLDocument(dlg.FileName))
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
                            List<MElements> lmeNume = new List<MElements>();
                            MElements dataNume = new MElements
                            {
                                Element = measureid.Numerator,
                                ElementData = sl.GetCellValueAsString(i + 2, 5).Trim(),
                                Eledate = dts
                            };
                            if (!gbackups.ContainsKey(measureid.Numerator))
                            {

                                lmeNume.Add(dataNume);
                                gbackups[measureid.Numerator] = lmeNume;
                            }
                            else
                            {
                                var meNume = gbackups[measureid.Numerator].FirstOrDefault(o => o.Eledate == dts);
                                if (meNume == null)
                                {
                                    gbackups[measureid.Numerator].Add(dataNume);
                                }
                            }
                        }
                        if (measureid.Denominator != "1")
                        {
                            List<MElements> lmeDeno = new List<MElements>();
                            MElements dataDeno = new MElements
                            {
                                Element = measureid.Denominator,
                                ElementData = sl.GetCellValueAsString(i + 2, 6).Trim(),
                                Eledate = dts
                            };
                            if (!gbackups.ContainsKey(measureid.Denominator))
                            {

                                lmeDeno.Add(dataDeno);
                                gbackups[measureid.Denominator] = lmeDeno;
                            }
                            else
                            {
                                var meDeno = gbackups[measureid.Denominator].FirstOrDefault(o => o.Eledate == dts);
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
                sl.Dispose();
            }
        }

        private void BT_EXPORT_ELEMENT(object sender, RoutedEventArgs e)
        {
            if (gbackups.Count <= 0 && gcollect.Count <= 0)
                return;

            //gcollect.Sort((x, y) => { return x.Element.CompareTo(y.Element); });

            var sortbacks = gbackups.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);

            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標收集存檔" + DateTime.Now.AddMonths(-1).ToString("yyyy-M") + ".xlsx";
            //string fname2 = @"\指標收集存檔總檔.xlsx";

            using (SLDocument sl = new SLDocument())
            {
                sl.SetColumnWidth(1, 15);
                sl.SetColumnWidth(2, 15);
                sl.SetCellValue(1, 1, "指標要素");
                for (int i = 0; i < 6; i++)
                    sl.SetCellValue(1, i + 2, DateTime.Now.AddMonths(-1 - i).ToString("yyyy/M"));
                int index = 0;
                foreach (var x in sortbacks)
                {
                    sl.SetCellValue(index + 2, 1, x.Key);

                    foreach (var y in x.Value)
                    {
                        for (int i = 0; i < 6; i++)
                        {
                            if (y.Eledate.Year == DateTime.Now.AddMonths(-1 - i).Year
                                && y.Eledate.Month == DateTime.Now.AddMonths(-1 - i).Month)
                            {
                                sl.SetCellValue(index + 2, i + 2, y.ElementData);
                            }
                        }
                    }

                    index++;
                }

                sl.SaveAs(fpath + fname);
                sl.Dispose();
            }
        }
    }
}
