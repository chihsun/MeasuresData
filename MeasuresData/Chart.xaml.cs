using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using LiveCharts;
using LiveCharts.Wpf;

namespace MeasuresData
{
    /// <summary>
    /// Chart.xaml 的互動邏輯
    /// </summary>
    public partial class Chart : Window
    {
        public Chart()
        {
            InitializeComponent();
            Values1 = new ChartValues<double>();
            Values2 = new ChartValues<double>();
            Values3 = new ChartValues<double>();
            Values4 = new ChartValues<double>();
            Values5 = new ChartValues<double>();
            Values6 = new ChartValues<double>();
            Labels = new List<string>();

        }

        public SeriesCollection SeriesCollection { get; set; }
        public ChartValues<double> Values1 { get; set; }
        public ChartValues<double> Values2 { get; set; }
        public ChartValues<double> Values3 { get; set; }
        public ChartValues<double> Values4 { get; set; }
        public ChartValues<double> Values5 { get; set; }
        public ChartValues<double> Values6 { get; set; }
        public string Values1Title { get; set; }
        //public string[] Labels { get; set; }
        public List<string> Labels { get; set; }
        public Func<double, string> YFormatter { get; set; }
        //public string Titles { get; set; }
        public List<MElement> GElements = new List<MElement>();
        //原始資料庫之指標資料定義
        public List<MMeasure> GMeasures = new List<MMeasure>();
        public SPC DataSet = new SPC();

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            List<string> comb_list = new List<string>();
            GMeasures.ForEach((o) =>
            {
                if (!comb_list.Contains(o.Group))
                    comb_list.Add(o.Group);
            });
            this.Combx0.ItemsSource = comb_list;
            this.Combx0.SelectedIndex = 0;
            List<string> spctype = new List<string>() { "Default", "U", "C", "P", "I_X", "I_MR", "nP", "Xbar_S", "Xbar_R" };
            Combx2.ItemsSource = spctype;
            Combx2.SelectedIndex = 0;
        }

        private void Combx0_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var lists = GMeasures.Where(o => o.Group == this.Combx0.SelectedValue.ToString()).Select(o => o.MeasureID).ToList();
            this.Combx1.ItemsSource = lists;
            this.Combx1.SelectedIndex = 0;
            this.Combx1.Focus();
        }

        private void Combx1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.Combx1.SelectedIndex < 0)
                return;
            var lb = GMeasures.FirstOrDefault(o => o.MeasureID == this.Combx1.SelectedValue.ToString());
            Lb_1.Content = string.Format("{0} - {1}", lb.MeasureID, lb.MeasureName);
            //Combx2.SelectedIndex = 0;
            var ctdata = GMeasures.FirstOrDefault(o => o.MeasureID == this.Combx1.SelectedValue.ToString()).Records; 
            var ctdatatake = ctdata.Reverse().ToDictionary(o => o.Key, o => o.Value);
            if (Ck_Sep.IsChecked == false)
            {
                List<string> exmonth = new List<string>();
                foreach (var x in ctdatatake)
                    exmonth.Add(x.Key);
                Combx4_Month_Ed.ItemsSource = Combx4_Month_ST.ItemsSource = exmonth;
                Combx4_Month_ST.SelectedIndex = 0;
                Combx4_Month_Ed.SelectedIndex = exmonth.Count - 1;
            }

            Combx2_SelectionChanged(sender, e);
        }

        private void Combx2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Paint();
        }
        private void Paint()
        {
            if (Combx2.SelectedIndex < 0 || Combx4_Month_Ed.SelectedIndex <= Combx4_Month_ST.SelectedIndex)
                return; 
            SPCtype spctype;
            if (Combx2.SelectedIndex == 0)
                spctype = SPCtype.P;
            else
                spctype = (SPCtype)Combx2.SelectedIndex;

            var ctdata = GMeasures.FirstOrDefault(o => o.MeasureID == this.Combx1.SelectedValue.ToString()).Records;
            if (ctdata == null || ctdata.Count <= 12)
                return;
            var ctdatatake = ctdata.Reverse().ToDictionary(o => o.Key, o => o.Value);
            SPC spcdatas;
            if (Ck_Sep.IsChecked == true)
            {
                if (Combx4_Month_ST.SelectedIndex < 3)
                    return;
                var data1 = new SPC(ctdatatake.Take(Combx4_Month_ST.SelectedIndex + 1).ToDictionary(o => o.Key, o => o.Value), spctype);
                var data2 = new SPC(ctdatatake.Skip(Combx4_Month_ST.SelectedIndex + 1).Take(Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex).ToDictionary(o => o.Key, o => o.Value), spctype);
                spcdatas = data1 + data2; 
                if ((Combx4_Month_Ed.Items.Count - Combx4_Month_Ed.SelectedIndex - 1) > 3)
                {
                    var data3 = new SPC(ctdatatake.Skip(Combx4_Month_Ed.SelectedIndex + 1).Take(Combx4_Month_Ed.Items.Count - Combx4_Month_Ed.SelectedIndex - 1).ToDictionary(o => o.Key, o => o.Value), spctype);
                    spcdatas += data3;
                }
            }
            else
            {
                if (Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex <= 3)
                    return;
                spcdatas = new SPC(ctdatatake.Skip(Combx4_Month_ST.SelectedIndex).Take(Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex + 1).ToDictionary(o => o.Key, o => o.Value), spctype);
            }
            double amply = 1;
            if (spcdatas.Type == SPCtype.P)
                amply = 100;
            int roundlevel = spcdatas.Average.Select(o => o * amply).Average() <= 0.1 ? 4 : 2;
            /*
            line1.Title = Combx1.SelectedValue.ToString();
            line1.Values = new ChartValues<double>(spcdatas.Measures.Select(o => Math.Round(o * amply, roundlevel)));
            line2.Values = new ChartValues<double>(spcdatas.Average.Select(o => Math.Round(o * amply, roundlevel)));
            line3.Values = new ChartValues<double>(spcdatas.UCL.Select(o => Math.Round(o * amply, roundlevel)));
            line4.Values = new ChartValues<double>(spcdatas.UUCL.Select(o => Math.Round(o * amply, roundlevel)));
            line5.Values = new ChartValues<double>(spcdatas.LCL.Select(o => Math.Round(o * amply, roundlevel)));
            line6.Values = new ChartValues<double>(spcdatas.LLCL.Select(o => Math.Round(o * amply, roundlevel)));
            */
            cartesianchart1.Series.Clear();

            cartesianchart1.Series.Add(new LineSeries
            {
                DataLabels = true,
                LineSmoothness = 0,
                PointGeometrySize = 10,
                PointForeground = (SolidColorBrush)(new BrushConverter().ConvertFrom("#222E31")),
                StrokeThickness = 4,
                StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 },
                Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#6BBA45")),
                Fill = Brushes.Transparent,
                Title = Combx1.SelectedValue.ToString(),
                Values = new ChartValues<double>(spcdatas.Measures.Select(o => Math.Round(o * amply, roundlevel)))
            });
            cartesianchart1.Series.Add(new LineSeries
            {
                LineSmoothness = 1,
                PointGeometry = null,
                StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 },
                Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#1C8FC5")),
                Fill = Brushes.Transparent,
                Title = "平均值",
                Values = new ChartValues<double>(spcdatas.Average.Select(o => Math.Round(o * amply, roundlevel)))
            });
            cartesianchart1.Series.Add(new LineSeries
            {
                LineSmoothness = 0,
                PointGeometry = null,
                StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF3333")),
                Fill = Brushes.Transparent,
                Title = "管制圖上限(2α)",
                Values = new ChartValues<double>(spcdatas.UCL.Select(o => Math.Round(o * amply, roundlevel)))
            });
            cartesianchart1.Series.Add(new LineSeries
            {
                LineSmoothness = 0,
                PointGeometry = null,
                StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFAD33")),
                Fill = Brushes.Transparent,
                Title = "管制圖上限(3α)",
                Values = new ChartValues<double>(spcdatas.UUCL.Select(o => Math.Round(o * amply, roundlevel)))
            });
            cartesianchart1.Series.Add(new LineSeries
            {
                LineSmoothness = 0,
                PointGeometry = null,
                StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF3333")),
                Fill = Brushes.Transparent,
                Title = "管制圖下限(2α)",
                Values = new ChartValues<double>(spcdatas.LCL.Select(o => Math.Round(o * amply, roundlevel)))
            });
            cartesianchart1.Series.Add(new LineSeries
            {
                LineSmoothness = 0,
                PointGeometry = null,
                StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFAD33")),
                Fill = Brushes.Transparent,
                Title = "管制圖下限(3α)",
                Values = new ChartValues<double>(spcdatas.LLCL.Select(o => Math.Round(o * amply, roundlevel)))
            });

            axisx.Sections.Clear();
            axisx.Labels = new List<string>(spcdatas.Title);
            if (Ck_Sep.IsChecked == true)
            {
                axisx.Sections.Add(new AxisSection()
                {
                    Value = Combx4_Month_ST.SelectedIndex + 1,
                    //Fill = Brushes.Transparent,
                    Fill = (SolidColorBrush)(new BrushConverter().ConvertFrom("#f9ffe6")),
                    Opacity = 7,
                    Stroke = Brushes.Transparent,
                    //StrokeThickness = 3,
                    //Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFFF33")),
                    //StrokeDashArray = new System.Windows.Media.DoubleCollection { 4 },
                    SectionWidth = Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex >= 3 ? Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex - 1 : 0
                });
                /*
                if (Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex >= 3)
                {
                    axisx.Sections.Add(new AxisSection()
                    {
                        Value = Combx4_Month_Ed.SelectedIndex,
                        Fill = Brushes.Transparent,
                        Opacity = 6,
                        StrokeThickness = 3,
                        Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFFF33")),
                        StrokeDashArray = new System.Windows.Media.DoubleCollection { 4 },
                        //SectionWidth = Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex
                    });
                }
                */
                /*
                axisx.Labels = new List<string>();
                foreach (var x in ctdatatake)
                    axisx.Labels.Add(x.Key);
                var data2 = new SPC(ctdatatake.Skip(Combx4_Month_ST.SelectedIndex + 1).Take(Combx4_Month_Ed.SelectedIndex - Combx4_Month_ST.SelectedIndex).ToDictionary(o => o.Key, o => o.Value), spctype);
                for (int i = 0; i < Combx4_Month_ST.SelectedIndex + 1; i++)
                {
                    data2.Title.Insert(0, string.Empty);
                    data2.Measures.Insert(0, double.NaN);
                    data2.Average.Insert(0, double.NaN);
                    data2.UCL.Insert(0, double.NaN);
                    data2.UUCL.Insert(0, double.NaN);
                    data2.LCL.Insert(0, double.NaN);
                    data2.LLCL.Insert(0, double.NaN);
                }
                cartesianchart1.Series.Add(new LineSeries
                {
                    DataLabels = true,
                    PointGeometrySize = 10,
                    PointForeground = (SolidColorBrush)(new BrushConverter().ConvertFrom("#222E31")),
                    StrokeThickness = 4,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 },
                    Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#6BBA45")),
                    Fill = Brushes.Transparent,
                    Title = null,
                    Values = new ChartValues<double>(data2.Measures.Select(o => Math.Round(o * amply, roundlevel)))
                });
                cartesianchart1.Series.Add(new LineSeries
                {
                    PointGeometry = null,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 },
                    Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#1C8FC5")),
                    Fill = Brushes.Transparent,
                    Title = null,
                    Values = new ChartValues<double>(data2.Average.Select(o => Math.Round(o * amply, roundlevel)))
                });
                cartesianchart1.Series.Add(new LineSeries
                {
                    PointGeometry = null,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                    Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF3333")),
                    Fill = Brushes.Transparent,
                    Values = new ChartValues<double>(data2.UCL.Select(o => Math.Round(o * amply, roundlevel)))
                });
                cartesianchart1.Series.Add(new LineSeries
                {
                    PointGeometry = null,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                    Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFAD33")),
                    Fill = Brushes.Transparent,
                    Values = new ChartValues<double>(data2.UUCL.Select(o => Math.Round(o * amply, roundlevel)))
                });
                cartesianchart1.Series.Add(new LineSeries
                {
                    PointGeometry = null,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                    Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF3333")),
                    Fill = Brushes.Transparent,
                    Values = new ChartValues<double>(data2.LCL.Select(o => Math.Round(o * amply, roundlevel)))
                });
                cartesianchart1.Series.Add(new LineSeries
                {
                    PointGeometry = null,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 1 },
                    Stroke = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFAD33")),
                    Fill = Brushes.Transparent,
                    Values = new ChartValues<double>(data2.LLCL.Select(o => Math.Round(o * amply, roundlevel)))
                });
                */
            }

            //YFormatter = value => value.ToString(roundlevel == 4 ? "F4" : "F2");

            cartesianchart1.AxisY[0].LabelFormatter = value => value.ToString(roundlevel == 4 ? "F4" : "F2");
            DataContext = this;
            cartesianchart1.Update();
            this.Combx1.Focus();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string fpath = Environment.CurrentDirectory + @"\管制圖";

            if (!System.IO.Directory.Exists(fpath))
            {
                System.IO.Directory.CreateDirectory(fpath);
            }
            var encoder = new PngBitmapEncoder();
            EncodeVisual(cartesianchart1, fpath + @"\SPC-(" + Combx1.SelectedValue.ToString() + DateTime.Now.ToString("yyyy-MM") + ").png", encoder);
        }
        private static void EncodeVisual(FrameworkElement visual, string fileName, BitmapEncoder encoder)
        {
            var bitmap = new RenderTargetBitmap((int)visual.ActualWidth, (int)visual.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            var frame = BitmapFrame.Create(bitmap);
            encoder.Frames.Add(frame);
            using (var stream = System.IO.File.Create(fileName)) encoder.Save(stream);
        }

        private void Ck_Sep_Checked(object sender, RoutedEventArgs e)
        {
            Paint();
        }
    }
}
