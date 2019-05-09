using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SpreadsheetLight;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;

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

        public void LoadFile(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return;
            
            try
            {
                SLDocument sl = new SLDocument(fname);
                //SLDocument sl2 = new SLDocument(fname, "工作表2");
                for (int i = 0; i < 500; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                        break;
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                        continue;
                    DiliveryData data = new DiliveryData();
                    data.Group = sl.GetCellValueAsString(i + 2, 1).Trim();
                    data.Depart = sl.GetCellValueAsString(i + 2, 2).Trim();
                    data.MeasureID = sl.GetCellValueAsString(i + 2, 3).Trim();
                    data.MeasureName = sl.GetCellValueAsString(i + 2, 5).Trim();
                    data.User = sl.GetCellValueAsString(i + 2, 7).Trim();

                    gdata.Add(data);

                    /*
                    if (PDatas.Count > 0)
                    {
                        bool duplicated = false;
                        foreach (var x in PDatas)
                        {
                            if (x.ID == data.ID && x.Name == data.Name)
                            {
                                x.Points += data.Points;
                                x.Detial += data.Detial;
                                duplicated = true;
                                break;
                            }
                        }
                        if (!duplicated)
                            PDatas.Add(data);
                    }
                    else
                        PDatas.Add(data);
                   */
                   
                }
                sl.CloseWithoutSaving();

                var newdata = gdata.GroupBy(o => o.Depart)
                    .ToDictionary(o => o.Key, o => o.ToList());

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
                    style.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
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
                    nsl.SetColumnStyle(5, style);

                    nsl.SetCellValue(1, 1, "指標群組");
                    nsl.SetCellValue(1, 2, "監視單位");
                    nsl.SetCellValue(1, 3, "指標要素");
                    nsl.SetCellValue(1, 4, "指標(要素)名稱");
                    nsl.SetCellValue(1, 5, "數值");
                    
                    for (int i = 0; i < x.Value.Count; i++)
                    {
                        nsl.SetCellValue(i + 2, 1, x.Value[i].Group);
                        nsl.SetCellValue(i + 2, 2, x.Value[i].Depart);
                        nsl.SetCellValue(i + 2, 3, x.Value[i].MeasureID);
                        nsl.SetCellValue(i + 2, 4, x.Value[i].MeasureName);
                    }
                    

                    string Refile = System.Environment.CurrentDirectory + @"\Group\" + x.Key + ".xlsx";
                    nsl.SaveAs(Refile);
                    nsl.Dispose();
                }
                MessageBox.Show("轉出成功 : " + gdata.Count.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Title = "選取資料檔";
            dlg.Filter = "xlsx files (*.*)|*.xlsx";
            if (dlg.ShowDialog() == true)
            {
                LoadFile(dlg.FileName);
            }
        }
        public void ExportData(string type)
        {
            if (type == "TCPI")
            {
                SLDocument nsl = new SLDocument();
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
                string time = DateTime.Now.ToString("yyy/M");
                if (gdata.Count > 0)
                {
                    int index = 2;
                    Random rand = new Random();
                    for (int i = 0; i < gdata.Count; i++)
                    {
                        if (gdata[i].Group != type)
                            continue;
                        nsl.SetCellValue(index, 1, time);
                        nsl.SetCellValue(index, 2, gdata[i].MeasureID);
                        nsl.SetCellValue(index, 3, gdata[i].MeasureName);
                        nsl.SetCellValue(index, 4, "月");
                        nsl.SetCellValue(index, 5, "55" + rand.Next(10, 200));
                        index++;
                    }
                }
                string Refile = System.Environment.CurrentDirectory + @"\Group\" + type + ".xlsx";
                nsl.SaveAs(Refile);
                nsl.Dispose();

            }
            else if (type == "評鑑持續")
            {
                if (!System.IO.File.Exists(System.Environment.CurrentDirectory + @"\201905.xlsx"))
                    return;
                SLDocument nsl = new SLDocument(System.Environment.CurrentDirectory + @"\201905.xlsx");
                Random rand = new Random();
                for (int i = 2; i < 94; i++)
                {
                    nsl.SetCellValue(i, 3, "55" + rand.Next(10, 200));
                }
                nsl.Save();
                nsl.Dispose();
                /*
                string time = DateTime.Now.ToString("yyy/MM");
                SLDocument nsl = new SLDocument();
                SLStyle style = nsl.CreateStyle();
                style.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
                style.Font.Bold = true;
                style.Font.FontSize = 12;
                style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 176, 240), System.Drawing.Color.Black);
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
                nsl.SetCellValue(1, 3, DateTime.Now.ToString("yyy/MM") + "(月)");
                nsl.SetCellValue(1, 4, DateTime.Now.AddMonths(-1).ToString("yyy/MM") + "(月)");
                nsl.SetCellValue(1, 5, DateTime.Now.AddMonths(-2).ToString("yyy/MM") + "(月)");
                nsl.SetCellValue(1, 6, DateTime.Now.AddMonths(-3).ToString("yyy/MM") + "(月)");
                nsl.SetCellValue(1, 7, DateTime.Now.AddMonths(-4).ToString("yyy/MM") + "(月)");
                nsl.SetCellValue(1, 8, DateTime.Now.AddMonths(-5).ToString("yyy/MM") + "(月)");
                
                if (gdata.Count > 0)
                {
                    var sortdata = gdata;
                    sortdata.Sort((x, y) => { return x.MeasureID.CompareTo(y.MeasureID); });
                    int index = 2;
                    Random rand = new Random();
                    for (int i = 0; i < sortdata.Count; i++)
                    {
                        if (sortdata[i].Group != type)
                            continue;
                        nsl.SetCellValue(index, 1, sortdata[i].MeasureID);
                        nsl.SetCellValue(index, 2, sortdata[i].MeasureName);
                        nsl.SetCellValue(index, 3, "55" + rand.Next(10, 200));
                        index++;
                    }
                }
                
                string Refile = System.Environment.CurrentDirectory + @"\Group\" + type + ".xlsx";
                nsl.SaveAs(Refile);
                nsl.Dispose();
                */
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ExportData("評鑑持續");
            MessageBox.Show("轉檔成功");
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
        public Dictionary<string, string> Value { get; set; }
        public DiliveryData()
        {
            Value = new Dictionary<string, string>();
        }

    }
}
