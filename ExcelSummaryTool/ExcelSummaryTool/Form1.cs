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
using OfficeOpenXml;
using Sunny.UI;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace ExcelSummaryTool
{
    public partial class Form1 : UIForm
    {
        private List<FileDetail> NoiseBeforeUnderFill_FileList = new List<FileDetail>();
        private List<FileDetail> NoiseAfterUnderFill_FileList = new List<FileDetail>();
        private List<FileDetail> SiganalBeforeUnderFill_FileList = new List<FileDetail>();
        private List<FileDetail> SiganalAfterUnderFill_FileList = new List<FileDetail>();
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("James");
        }
        #region Functions
        private object[,,] ReadExcelData(string path)
        {
            object[,,] excelData = null;
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                int sheetCount = package.Workbook.Worksheets.Count;
                int maxRow = 0;
                int maxCol = 0;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    if (ws.Dimension != null)
                    {
                        if (ws.Dimension.Rows > maxRow) maxRow = ws.Dimension.Rows;
                        if (ws.Dimension.Columns > maxCol) maxCol = ws.Dimension.Columns;
                    }
                }
                excelData = new object[sheetCount, maxRow, maxCol];
                for (int s = 0; s < sheetCount; s++)
                {
                    var ws = package.Workbook.Worksheets[s];
                    int rows = ws.Dimension?.Rows ?? 0;
                    int cols = ws.Dimension?.Columns ?? 0;

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            excelData[s, r - 1, c - 1] = ws.Cells[r, c].Value;
                        }
                    }
                }
            }
            return excelData;
        }
        /// <summary>
        /// 取得指定頁面,指定欄位的資料陣列
        /// </summary>
        /// <param name="data_array"></param>
        /// <param name="sheet_index"></param>
        /// <param name="column_index"></param>
        /// <returns></returns>
        private object[] GetSpecifyDataArray(object[,,] data_array,int sheet_index,int column_index)
        {
            int rowCount = data_array.GetLength(1);
            object[] specify_data = new object[rowCount];
            for(int i = 0; i < rowCount; i++)
            {
                specify_data[i] = data_array[sheet_index, i, column_index];
            }
            return specify_data;
        }
        private void Tab1_Result_Table()
        {
            #region 參數
            int sheet_index = Convert.ToInt32(Tab1_Sheet_tb.Text);
            int column_index = ExcelColumnToNumber(Tab1_Column_tb.Text);
            #endregion
            #region Get BeforeUnderfill 
            List<DataDetail> BUF_data_list = new List<DataDetail>();
            foreach (var file in NoiseBeforeUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail(); 
                object[] BUF_data_array;
                BUF_data_array = GetSpecifyDataArray((object[,,])file.Data,sheet_index,column_index);
                tmp.Data = BUF_data_array;
                tmp.SN = file.SN;
                BUF_data_list.Add(tmp);
            }
            #endregion

            #region Get AfterUnderfill 
            List<DataDetail> AUF_data_list = new List<DataDetail>();
            foreach (var file in NoiseAfterUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] AUF_data_array;
                AUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index, column_index);
                tmp.Data = AUF_data_array;
                tmp.SN = file.SN;
                AUF_data_list.Add(tmp);
            }
            #endregion

            #region Calculation
            var dict1 = BUF_data_list.ToDictionary(d => d.SN);
            var dict2 = AUF_data_list.ToDictionary(d => d.SN);

            List<DataDetail> diffList = new List<DataDetail>();

            foreach (var sn in dict1.Keys)
            {
                if (dict2.ContainsKey(sn))
                {
                    var data1 = dict1[sn].Data;
                    var data2 = dict2[sn].Data;

                    // 確保長度一致
                    int len = Math.Min(data1.Length, data2.Length);
                    object[] diff = new object[len];

                    for (int i = 0; i < len; i++)
                    {
                        double val1 = data1[i] != null ? Convert.ToDouble(data1[i]) : 0;
                        double val2 = data2[i] != null ? Convert.ToDouble(data2[i]) : 0;
                        diff[i] = val1 - val2;
                    }

                    diffList.Add(new DataDetail
                    {
                        SN = sn,
                        Data = diff
                    });
                }
            }
            #endregion

            #region Merge to ExcelSheet
            
            int buf_rowCount = BUF_data_list.Max(d => d.Data.Length); // 取最高的列數
            int buf_colCount = BUF_data_list.Count;                  // 每個 Data 一欄

            object[,] buf2D = new object[buf_rowCount, buf_colCount];

            for (int c = 0; c < buf_colCount; c++)
            {
                var data = BUF_data_list[c].Data;
                for (int r = 0; r < data.Length; r++)
                {
                    buf2D[r, c] = data[r]; // row = Data 的元素索引, column = DataDetail 的索引
                }
            }
            int auf_rowCount = AUF_data_list.Max(d => d.Data.Length); // 取最高的列數
            int auf_colCount = AUF_data_list.Count;                  // 每個 Data 一欄

            object[,] auf2D = new object[auf_rowCount, auf_colCount];

            for (int c = 0; c < auf_colCount; c++)
            {
                var data = AUF_data_list[c].Data;
                for (int r = 0; r < data.Length; r++)
                {
                    auf2D[r, c] = data[r]; // row = Data 的元素索引, column = DataDetail 的索引
                }
            }
            int diff_rowCount = diffList.Max(d => d.Data.Length); // 取最高的列數
            int diff_colCount = diffList.Count;                  // 每個 Data 一欄

            object[,] diff2D = new object[diff_rowCount, diff_colCount];

            for (int c = 0; c < diff_colCount; c++)
            {
                var data = diffList[c].Data;
                for (int r = 0; r < data.Length; r++)
                {
                    diff2D[r, c] = data[r]; // row = Data 的元素索引, column = DataDetail 的索引
                }
            }
            int totalRowCount = Math.Max(buf_rowCount, Math.Max(auf_rowCount, diff_rowCount));
            int totalColCount = buf_colCount + auf_colCount + diff_colCount;

            object[,] result_array = new object[totalRowCount, totalColCount];
            // 塞 BUF
            for (int r = 0; r < buf_rowCount; r++)
            {
                for (int c = 0; c < buf_colCount; c++)
                {
                    result_array[r, c] = buf2D[r, c];
                }
            }

            // 塞 AUF
            for (int r = 0; r < auf_rowCount; r++)
            {
                for (int c = 0; c < auf_colCount; c++)
                {
                    result_array[r, buf_colCount + c] = auf2D[r, c];
                }
            }

            // 塞 diff
            for (int r = 0; r < diff_rowCount; r++)
            {
                for (int c = 0; c < diff_colCount; c++)
                {
                    result_array[r, buf_colCount + auf_colCount + c] = diff2D[r, c];
                }
            }
            #endregion
        }
        private void CreateExcelTable()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Summary");
                
                var fileInfo = new FileInfo("Summary.xlsx");
                package.SaveAs(fileInfo);
            }
        }
        #endregion
        #region UI_Events
        private async void SelectFolder_Click(object sender, EventArgs e)
        {
            UITextBox uITextBox = (UITextBox)sender;
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = Properties.Settings.Default.LastFolderPath;
            folderBrowserDialog.Description = "請選擇要輸入的資料夾";
            folderBrowserDialog.ShowNewFolderButton = false;
            DialogResult result = folderBrowserDialog.ShowDialog();
            string[] files;
            if (result == DialogResult.OK)
            {
                Properties.Settings.Default.LastFolderPath = folderBrowserDialog.SelectedPath;
                Properties.Settings.Default.Save();
                uITextBox.Text = folderBrowserDialog.SelectedPath;
                string[] allFiles = Directory.GetFiles(folderBrowserDialog.SelectedPath);
                files = allFiles
                        .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                        && !Path.GetFileName(f).StartsWith("~$"))
                        .ToArray();
                List<FileDetail> file_list = new List<FileDetail>();
                for (int i = 0; i < files.Length; i++)
                {
                    FileDetail fileDetail = new FileDetail();
                    fileDetail.FilePath = folderBrowserDialog.SelectedPath;
                    fileDetail.FIleName = files[i];
                    fileDetail.Data = await Task.Run(() => ReadExcelData(files[i]));
                    //fileDetail.SN = 每個檔案的SN序號
                    file_list.Add(fileDetail);
                }
                Console.WriteLine($"TextBox:{uITextBox.Name},Path:{uITextBox.Text},ReadFileListData Done");
                switch (uITextBox.Name)
                {
                    case "uiTextBox1":
                        NoiseBeforeUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox2":
                        NoiseAfterUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox3":
                        SiganalBeforeUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox4":
                        SiganalAfterUnderFill_FileList = file_list;
                        break;
                }
            }
        }
        private void Start_Bt_Click(object sender, EventArgs e)
        {

        }
        #endregion
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
            return sum-1;
        }
        #endregion
    }
}
