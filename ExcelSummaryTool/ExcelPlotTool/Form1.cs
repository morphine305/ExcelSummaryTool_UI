using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using Sunny.UI;

namespace ExcelPlotTool
{
    public partial class Form1 : UIForm
    {
        List<string> File_list;
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("James");
        }
        private async void SelectFolder_Click(object sender, EventArgs e)
        {
            UITextBox uITextBox = (UITextBox)sender;
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = Properties.Settings.Default.LastFolderPath;
            folderBrowserDialog.Description = "請選擇要輸入的資料夾";
            folderBrowserDialog.ShowNewFolderButton = false;
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                Properties.Settings.Default.LastFolderPath = folderBrowserDialog.SelectedPath;
                Properties.Settings.Default.Save();
                uITextBox.Text = folderBrowserDialog.SelectedPath;
                string[] allFiles = Directory.GetFiles(folderBrowserDialog.SelectedPath);
                File_list = allFiles
                        .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                        && !Path.GetFileName(f).StartsWith("~$"))
                        .ToList();
            }
        }
        public static ExcelScatterChart CreateXYScatterChart_AcrossSheet(
        ExcelWorksheet ws1, ExcelWorksheet ws2, ExcelWorksheet ws3, ExcelWorksheet ws4,
        string chartName,
        int xCol,
        int yCol,
        int startRow,
        int endRow, double minvalue = 60, double maxvalue = 63.5,
        int posRow = 1,
        int posCol = 3,
        int width = 600,
        int height = 400,
        string seriesHeader = "Series")
        {
            // 建立圖表
            var chart = ws4.Drawings.AddChart(chartName, eChartType.XYScatter) as ExcelScatterChart;

            chart.Title.Text = chartName;

            // 建立 Y, X 資料範圍
            var xRange = ws1.Cells[startRow, xCol, endRow, xCol];
            var yRange = ws2.Cells[startRow, yCol, endRow, yCol];

            var series = chart.Series.Add(yRange, xRange);
            series.Header = seriesHeader;

            chart.SetPosition(posRow, 0, posCol, 0);
            chart.SetSize(width, height);
            chart.XAxis.MinValue = minvalue; chart.XAxis.MaxValue = maxvalue;

            return chart; // 回傳圖表物件
        }
        public List<ChartSeries> CreateSeriesList(ExcelWorksheet ws,int rowcount,int start_col,int end_col)
        {
            List<ChartSeries> series_list = new List<ChartSeries>();
           
            for(int i = 2; i < rowcount; i++)
            {
                ChartSeries series1 = new ChartSeries();
                series1.X = ws.Cells[1, start_col, 1, end_col];
                series1.Y = ws.Cells[i, start_col, i, end_col];
                series1.Header = ws.Cells[i,1].Value.ToString();
                series_list.Add(series1);
            }
            return series_list;
            
        }
        public static ExcelChart CreateXYSmoothLineChart(ExcelWorksheet ws2,ExcelWorksheet ws4, 
            string chartName, List<ChartSeries> seriesList,
            int posRow = 1,
            int posCol = 15,
            int width = 800,
            int height = 400)
        {
            var chart = ws4.Drawings.AddChart(chartName, eChartType.XYScatterSmooth);
            chart.Title.Text = chartName;
            chart.SetPosition(posRow, 0, posCol, 0);
            chart.SetSize(width, height);
            foreach (var s in seriesList)
            {
                var series = chart.Series.Add(s.Y, s.X);
                series.Header = s.Header;
            }
            return chart;
        }
        #region Excel 小工具
        /// <summary>
        /// 將 Excel 欄位字母轉成欄位數字（A=1, B=2 ... AA=27）
        /// </summary>
        public static int ExcelColumnToNumber(string column)
        {
            if (string.IsNullOrEmpty(column)) return 0;

            column = column.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < column.Length; i++)
            {
                sum *= 26;
                sum += (column[i] - 'A' + 1);
            }
            return sum;
        }
        #endregion

        private void start_bt_Click(object sender, EventArgs e)
        {
            foreach (var filePath in File_list)
            {
                var file = new FileInfo(filePath);
                if (!file.Exists) continue;
                using (var package = new ExcelPackage(file))
                {
                    var sheet1 = package.Workbook.Worksheets[0];
                    var sheet2 = package.Workbook.Worksheets[1];
                    var sheet3 = package.Workbook.Worksheets[2];
                    var wsChart = package.Workbook.Worksheets.Add("Chart");
                    string a2 = sheet3.Cells[2, 1].Value.ToString(); //sheet 2 a2 標題
                    #region chamber SNR_B比chamber SNR_diff
                    int rowcount = sheet1.Dimension?.End.Row ?? 0;
                    string[] chamber_title = chartTitle("SNR vs SNR_Chamber_Diff_", a2);

                    int[] xC = chart_array(ExcelColumnToNumber("ACN"));
                    int[] yC = chart_array(ExcelColumnToNumber("OC"));
                    for (int x = 0; x < xC.Length; x++)
                    {
                        var chart = CreateXYScatterChart_AcrossSheet(sheet1, sheet2, sheet3, wsChart, chamber_title[x], xC[x], yC[x], 2, rowcount);
                    }
                    #endregion
                    #region  Radar demo SNR_B比Radar demo SNR_diff
                    string[] PCapp_title = new string[] { $"SNR  vs SNR_RadarDemo_Diff_50_{a2}",$"SNR  vs SNR_RadarDemo_Diff_0_{a2}",$"SNR  vs SNR_RadarDemo_Diff_-50_{a2}" };
                    string[] PC_xC = new string[] {"AMA", "ANY", "APW" };
                    string[] PC_yC = new string[] { "SY", "UW", "WU" };

                    xC = ColumnToNum_Array(PC_xC);
                    yC = ColumnToNum_Array(PC_yC);
                    for(int x = 0; x < xC.Length; x++)
                    {
                        var chart = CreateXYScatterChart_AcrossSheet(sheet1, sheet2, sheet3, wsChart, PCapp_title[x], xC[x], yC[x],2,rowcount,minvalue:59,maxvalue:67,posRow:1,posCol:10);
                    }
                    #endregion
                    #region chamber SNR_diff(+60~60deg曲線圖)
                    string[] start_col_str = new string[] {"NX","SO" };
                    string[] end_col_str = new string[] { "SN", "XE" };
                    string[] chartName = new string[] { $"cahmber SNR_diff_{a2}",$"Radar demo SNR_diff_{a2}" };
                    for(int i = 0; i < start_col_str.Length; i++)
                    {
                        int start_col = ExcelColumnToNumber(start_col_str[i]);
                        int end_col = ExcelColumnToNumber(end_col_str[i]);
                        List<ChartSeries> series = CreateSeriesList(sheet2, rowcount, start_col, end_col);
                        var chart = CreateXYSmoothLineChart(sheet2, wsChart, chartName[i],series);
                    }
                    

                    #endregion
                    package.Save();
                }
            }
            MessageBox.Show("Done");
        }
        private string[] chartTitle(string name, string a2name)
        {
            string[] num = new string[39];
            int c = 55;
            for (int i = 0; i < 11; i++)
            {
                num[i] = $"{name}{c.ToString()}_{a2name}";
                c--;
            }
            int x = 11;
            c = c - 4;
            while (c >= -45)
            {
                num[x++] = $"{name}{c.ToString()}_{a2name}";
                c -= 5;
            }
            c = c + 4;
            for (int i = x; i < 39; i++)
            {
                num[i] = $"{name}{c.ToString()}_{a2name}";
                c--;
            }
            return num;
        }
        private int[] chart_array(int start)
        {
            int[] num = new int[39];
            for (int i = 0; i < 11; i++)
            {
                num[i] = start;
                start++;
            }
            start += 4;
            int j = 11;
            for (int i = 0; i < 18; i++)
            {
                num[j] = start;
                start += 5;
                j++;
            }
            start -= 4;
            for (int i = 0; i < 10; i++)
            {
                num[j + i] = start++;
            }
            return num;
        }
        private int[] ColumnToNum_Array(string[] cols)
        {
            int[] col_index = new int[cols.Length];
            for (int i = 0; i < cols.Length; i++)
            {
                col_index[i] = ExcelColumnToNumber(cols[i]);
            }
            return col_index;
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public class ChartSeries
        {
            public ExcelRange Y { get; set; }
            public ExcelRange X { get; set; }
            public string Header { get; set; }
        }
    }
}
