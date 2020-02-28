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
            /*
            var ctdata = GMeasures.FirstOrDefault(o => o.MeasureID == "HA02-01").Records;
            Titles = GMeasures.FirstOrDefault(o => o.MeasureID == "HA02-01").MeasureName;
            if (ctdata == null || ctdata.Count <= 12)
                return;
            var spcdatas = new SPC(ctdata);
            Values1 = new ChartValues<double>();
            Values1.AddRange(spcdatas.Measures);
            Values2 = new ChartValues<double>();
            Values2.AddRange(spcdatas.Average);
            Values3 = new ChartValues<double>();
            Values3.AddRange(spcdatas.UCL);
            Values4 = new ChartValues<double>();
            Values4.AddRange(spcdatas.UUCL);
            Values5 = new ChartValues<double>();
            Values5.AddRange(spcdatas.LCL);
            Values6 = new ChartValues<double>();
            Values6.AddRange(spcdatas.LLCL);
            Labels = new List<string>();
            Labels.AddRange(spcdatas.Title);
            Values1Title = "HA02-01";
            */

            /*
            ChartValues<double> ctva = new ChartValues<double>();
            int i = 0;
            foreach (var x in ctdata)
            {
                if (i >= 12)
                    break;
                ctva.Insert(0, x.Value[1] == 0 ? 0 : Math.Round(1000 * x.Value[0] / x.Value[1], 2));
                Labels.Insert(0, x.Key);
                i++;
            }
            
            SeriesCollection = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "HA01-01",
                    Values = ctva,
                    Fill = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
                },
                new LineSeries
                {
                    Title = "Series 2",
                    Values = new ChartValues<double> { 6, 7, 3, 4 ,6 },
                    PointGeometry = null
                },
                new LineSeries
                {
                    Title = "Series 3",
                    Values = new ChartValues<double> { 4,2,7,2,7 },
                    PointGeometry = DefaultGeometries.Square,
                    PointGeometrySize = 15
                }
            };
            */
            /*
            Labels = new[] { "Jan", "Feb", "Mar", "Apr", "May" };
            //YFormatter = value => value.ToString("C");

            //modifying the series collection will animate and update the chart
            SeriesCollection.Add(new LineSeries
            {
                Title = "Series 4",
                Values = new ChartValues<double> { 5, 3, 2, 4 },
                LineSmoothness = 0, //0: straight lines, 1: really smooth lines
                PointGeometry = Geometry.Parse("m 25 70.36218 20 -28 -20 22 -8 -6 z"),
                PointGeometrySize = 50,
                PointForeground = Brushes.Gray
            });

            //modifying any series values will also animate and update the chart
            SeriesCollection[3].Values.Add(5d);
            DataContext = this;
            */
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
            Combx2_SelectionChanged(sender, e);
        }

        private void Combx2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Combx2.SelectedIndex < 0)
                return;
            SPCtype spctype;
            if (Combx2.SelectedIndex == 0)
                spctype = SPCtype.P;
            else
                spctype = (SPCtype)Combx2.SelectedIndex;

            var ctdata = GMeasures.FirstOrDefault(o => o.MeasureID == this.Combx1.SelectedValue.ToString()).Records;
            if (ctdata == null || ctdata.Count <= 12)
                return;
            var spcdatas = new SPC(ctdata.Take(12).Reverse().ToDictionary(o => o.Key, o => o.Value), spctype);
            double amply = 1;
            if (spctype == SPCtype.P)
                amply = 100;
            int roundlevel = spcdatas.Average.Select(o => o * amply).Average() <= 0.1 ? 4 : 2;
            line1.Title = Combx1.SelectedValue.ToString();
            line1.Values = new ChartValues<double>(spcdatas.Measures.Select(o => Math.Round(o * amply, roundlevel)));
            line2.Values = new ChartValues<double>(spcdatas.Average.Select(o => Math.Round(o * amply, roundlevel)));
            line3.Values = new ChartValues<double>(spcdatas.UCL.Select(o => Math.Round(o * amply, roundlevel)));
            line4.Values = new ChartValues<double>(spcdatas.UUCL.Select(o => Math.Round(o * amply, roundlevel)));
            line5.Values = new ChartValues<double>(spcdatas.LCL.Select(o => Math.Round(o * amply, roundlevel)));
            line6.Values = new ChartValues<double>(spcdatas.LLCL.Select(o => Math.Round(o * amply, roundlevel)));

            //axisx.Title = GMeasures.FirstOrDefault(o => o.MeasureID == Combx1.SelectedValue.ToString()).MeasureName;
            axisx.Labels = new List<string>(spcdatas.Title);
            //if (spctype == SPCtype.U || spctype == SPCtype.P)
                //YFormatter = value => value.ToString("P", System.Globalization.CultureInfo.InvariantCulture);
            YFormatter = value => value.ToString(roundlevel == 4 ? "F4" : "F2");


            DataContext = this;
            cartesianchart1.AxisY[0].LabelFormatter = value => value.ToString(roundlevel == 4 ? "F4" : "F2");
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
    }
}
