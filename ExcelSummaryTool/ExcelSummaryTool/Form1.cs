using System;
using System.Collections;
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
        private List<FileDetail> RangeProBeforeUnderFill_FileList = new List<FileDetail>();
        private List<FileDetail> RangeProAfterUnderFill_FileList = new List<FileDetail>();

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
        /// 取得指定頁面,指定列,指定欄的資料陣列
        /// </summary>
        /// <param name="data_array"></param>
        /// <param name="sheet_index"></param>
        /// <param name="start_row_index"></param>
        /// <param name="start_column_index"></param>
        /// <param name="end_row_index"></param>
        /// <param name="end_column_index"></param>
        /// <returns></returns>
        private object[,] GetSelectDataArray(object[,,] data_array, int sheet_index, int start_row_index,int start_column_index, int end_row_index, int end_column_index)
        {
            int rowcount = end_row_index - start_row_index +1;
            int columncount = end_column_index - start_column_index +1;
            object[,] select_data = new object[rowcount,columncount];
            for (int r = 0; r < rowcount; r++)
            {
                for (int c = 0; c < columncount; c++)
                {
                    select_data[r, c] = data_array[sheet_index, start_row_index + r, start_column_index + c];
                }
            }
            return select_data;
        }
        private object[] AvgNoise_Array(object[,] data_array)
        {
            int rows = data_array.GetLength(0);
            int cols = data_array.GetLength(1);
            object[] avg = new object[data_array.GetLength(1)];
            for (int j = 0; j < cols; j++)
            {
                double total = 0;
                for (int i = 0; i < rows; i++)
                {
                    if (data_array[i, j] != null)
                        total += Convert.ToDouble(data_array[i, j]);
                }
                avg[j] = total / rows;
            }
            return avg;
        }
        private object AvgNoise(object[] data_array)
        {
            double total = 0;
            for (int i = 0; i < data_array.Length; i++)
            {
                if(data_array[i] != null)
                {
                    total += Convert.ToDouble(data_array[i]);
                }
            }
            return total / data_array.Length;
        }
        /// <summary>
        /// 取得指定頁面,指定欄位的資料陣列
        /// </summary>
        /// <param name="data_array"></param>
        /// <param name="sheet_index"></param>
        /// <param name="column_index"></param>
        /// <returns></returns>
        private object[] GetSpecifyDataArray(object[,,] data_array, int sheet_index, int column_index)
        {
            int rowCount = data_array.GetLength(1);
            object[] specify_data = new object[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                specify_data[i] = data_array[sheet_index, i, column_index];
            }
            return specify_data;
        }
        private void GetTableData(out List<Table> tabledata)
        {
            #region 參數
            int sheet_index_signal = Convert.ToInt32(Tab1_Sheet_tb.Text) - 1;
            int column_index_signal = ExcelColumnToNumber(Tab1_Column_tb.Text);
            int sheet_index_noise = Convert.ToInt32(Tab2_Sheet_tb.Text) - 1;
            int column_index_noise = ExcelColumnToNumber(Tab2_Column_tb.Text);
            #endregion
            tabledata = new List<Table>(); 
            #region Get BeforeUnderfill 
            List<DataDetail> BUF_Signal_list = new List<DataDetail>();
            List<DataDetail> BUF_Noise_list = new List<DataDetail>();
            foreach (var file in SiganalBeforeUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] BUF_data_array;
                BUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index_signal, column_index_signal);
                tmp.Data = BUF_data_array;
                tmp.SN = file.SN;
                BUF_Signal_list.Add(tmp);
            }
            foreach (var file in NoiseBeforeUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] BUF_data_array;
                BUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index_noise, column_index_noise);
                tmp.Data = BUF_data_array;
                tmp.SN = file.SN;
                BUF_Noise_list.Add(tmp);
            }
            #endregion
            #region Get AfterUnderfill 
            List<DataDetail> AUF_Signal_list = new List<DataDetail>();
            List<DataDetail> AUF_Noise_list = new List<DataDetail>();
            foreach (var file in SiganalAfterUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] AUF_data_array;
                AUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index_signal, column_index_signal);
                tmp.Data = AUF_data_array;
                tmp.SN = file.SN;
                AUF_Signal_list.Add(tmp);
            }
            foreach (var file in NoiseAfterUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] AUF_data_array;
                AUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index_noise, column_index_noise);
                tmp.Data = AUF_data_array;
                tmp.SN = file.SN;
                AUF_Noise_list.Add(tmp);
            }
            #endregion
            int count = BUF_Signal_list.Count;
            for(int i = 0; i < count; i++)
            {
                Table table = new Table();
                table.Name = BUF_Signal_list[i].SN;
                table.Signal_Before = BUF_Signal_list[i].Data;
                var target = AUF_Signal_list.FirstOrDefault(x => x.SN == table.Name);
                if(target != null)
                {
                    table.Signal_After = target.Data;
                }
                else
                {
                    table.Signal_After = null;
                }
                table.Noise_Before = BUF_Noise_list[i].Data;
                target = AUF_Noise_list.FirstOrDefault(x => x.SN == table.Name);
                if (target != null)
                {
                    table.Noise_After = target.Data;
                }
                else
                {
                    table.Noise_After = null;
                }
                tabledata.Add(table);
            }
        }
        private void GetRangeTable()
        {
            #region 參數
            int sheet_index = Convert.ToInt32(TB3_sheet_tb.Text);
            int start_col_index = ExcelColumnToNumber(TB3_Start_Column_tb.Text);
            int end_col_index = ExcelColumnToNumber(TB3_End_Column_tb.Text);
            int start_row_index = Convert.ToInt32(TB3_Start_Row_tb.Text);
            int end_row_index = Convert.ToInt32(TB3_End_Row_tb.Text);
            #endregion
            #region Get BeforeUnderfill 
            List<DataDetail> BUF_Range_list = new List<DataDetail>();
            List<DataDetail> AUF_Range_list = new List<DataDetail>();
            foreach (var file in RangeProBeforeUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[,] BUF_data_array;
                BUF_data_array = GetSelectDataArray((object[,,])file.Data,sheet_index,start_row_index,start_col_index,end_row_index,end_col_index);
                tmp.Data = AvgNoise_Array(BUF_data_array);
                tmp.SN = file.SN;
                BUF_Range_list.Add(tmp);
            }
            #endregion
            #region Get AfterUnderfill 
            foreach (var file in RangeProAfterUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[,] AUF_data_array;
                AUF_data_array = GetSelectDataArray((object[,,])file.Data, sheet_index, start_row_index, start_col_index, end_row_index, end_col_index);
                tmp.Data = AvgNoise_Array(AUF_data_array);
                tmp.SN = file.SN;
                AUF_Range_list.Add(tmp);
            }
            #endregion
        }
        private List<Table> GetSNR(List<Table> table_list)
        {
            foreach(var table in table_list)
            {
                object[] snr_before = new object[table.Signal_Before.Length];
                for(int i = 1; i < table.Signal_Before.Length; i++)
                {
                    double noise = Convert.ToDouble(table.Noise_Before[1]);
                    snr_before[i] = Convert.ToDouble(table.Signal_Before[i]) - noise;
                }
                table.SNR_Before = snr_before;

                if (table.Signal_After == null)
                {
                    table.SNR_After = null;
                }
                else
                {
                    object[] snr_after = new object[table.Signal_After.Length];
                    for (int i = 1; i < table.Signal_After.Length; i++)
                    {
                        double noise = Convert.ToDouble(table.Noise_After[1]);
                        snr_after[i] = Convert.ToDouble(table.Signal_After[i]) - noise;
                    }
                    table.SNR_After = snr_after;
                }
            }
            return table_list;
        }
        private void Tab1_Result_Table(out object[,] Tab1_array, out List<SNRData> Signal_list)
        {
            #region 參數
            int sheet_index = Convert.ToInt32(Tab1_Sheet_tb.Text) - 1;
            int column_index = ExcelColumnToNumber(Tab1_Column_tb.Text);
            #endregion
            #region Get BeforeUnderfill 
            List<DataDetail> BUF_data_list = new List<DataDetail>();
            List<TempData> BFS_data = new List<TempData>();
            foreach (var file in SiganalBeforeUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] BUF_data_array;
                BUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index, column_index);
                tmp.Data = BUF_data_array;
                tmp.SN = file.SN;
                BUF_data_list.Add(tmp);
                #region Prepare for SNR data
                TempData temp = new TempData();
                temp.objects = BUF_data_array;
                temp.SN = file.SN;
                BFS_data.Add(temp);
                #endregion
            }
            #endregion

            #region Get AfterUnderfill 
            List<DataDetail> AUF_data_list = new List<DataDetail>();
            List<TempData> AFS_data = new List<TempData>();
            foreach (var file in SiganalAfterUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] AUF_data_array;
                AUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index, column_index);
                tmp.Data = AUF_data_array;
                tmp.SN = file.SN;
                AUF_data_list.Add(tmp);
                #region Prepare for SNR data
                TempData temp = new TempData();
                temp.objects = AUF_data_array;
                temp.SN = file.SN;
                AFS_data.Add(temp);
                #endregion
            }
            #endregion
            #region Calculation
            var dictBUF = BUF_data_list.ToDictionary(d => d.SN);
            var dictAUF = AUF_data_list.ToDictionary(d => d.SN);
            List<DataDetail> diffList = new List<DataDetail>();

            foreach (var sn in dictBUF.Keys)
            {
                if (dictAUF.ContainsKey(sn))
                {
                    var dataBUF = dictBUF[sn].Data;
                    var dataAUF = dictAUF[sn].Data;
                    int len = Math.Min(dataBUF.Length, dataAUF.Length);
                    object[] diff = new object[len];
                    diff[0] = dataBUF[0];

                    for (int i = 1; i < len; i++)
                    {
                        double valBUF = dataBUF[i] != null ? Convert.ToDouble(dataBUF[i]) : 0;
                        double valAUF = dataAUF[i] != null ? Convert.ToDouble(dataAUF[i]) : 0;
                        diff[i] = valBUF - valAUF;
                    }

                    diffList.Add(new DataDetail { SN = sn, Data = diff });
                }
            }
            #endregion
            #region Prepare for SNR calculation
            Signal_list = new List<SNRData>();
            Signal_list = (from a in BFS_data
                           join b in AFS_data on a.SN equals b.SN into gj
                           from subB in gj.DefaultIfEmpty() // 沒對應就 subB 為 null
                           select new SNRData
                           {
                               SN = a.SN,
                               BeforeData = a.objects,
                               AfterData = subB?.objects ?? Array.Empty<object>()  // 找不到就 null
                           }).ToList();
            #endregion
            #region Merge by SN
            // X 軸欄位 = BUF.Count + 2空欄 + AUF.Count + 2空欄 + Diff.Count
            int colCount = BUF_data_list.Count + 2 + AUF_data_list.Count + 2 + diffList.Count;

            // Y 軸 = 資料最大長度
            int rowCount = Math.Max(
                BUF_data_list.Count > 0 ? BUF_data_list.Max(d => d.Data.Length) : 0,
                Math.Max(
                    AUF_data_list.Count > 0 ? AUF_data_list.Max(d => d.Data.Length) : 0,
                    diffList.Count > 0 ? diffList.Max(d => d.Data.Length) : 0
                )
            );


            object[,] result_array = new object[rowCount, colCount];

            // 塞 BUF (第一段欄)
            for (int c = 0; c < BUF_data_list.Count; c++)
            {
                var data = BUF_data_list[c].Data;
                for (int r = 0; r < data.Length; r++)
                    result_array[r, c] = data[r];
            }

            // 塞 AUF (空 2 欄後)
            int aufStart = BUF_data_list.Count + 2;
            for (int c = 0; c < AUF_data_list.Count; c++)
            {
                var data = AUF_data_list[c].Data;
                for (int r = 0; r < data.Length; r++)
                    result_array[r, aufStart + c] = data[r];
            }

            // 塞 Diff (空 2 + BUF長 + 2 + AUF長)
            int diffStart = BUF_data_list.Count + 2 + AUF_data_list.Count + 2;
            for (int c = 0; c < diffList.Count; c++)
            {
                var data = diffList[c].Data;
                for (int r = 0; r < data.Length; r++)
                {
                    if (data[r] is string)
                    {
                        result_array[r, diffStart + c] = data[r]; // 保留原字串
                    }
                    else
                    {
                        double val = Convert.ToDouble(data[r]);
                        result_array[r, diffStart + c] = (object)val;
                    }
                }
            }
            #endregion
            Tab1_array = result_array;
        }
        private void Tab2_Result_Table(out object[,] Tab2_array, out List<SNRData> Noise_list)
        {
            #region 參數
            int sheet_index = Convert.ToInt32(Tab2_Sheet_tb.Text) - 1;
            int column_index = ExcelColumnToNumber(Tab2_Column_tb.Text);
            #endregion
            #region Get BeforeUnderfill 
            List<DataDetail> BUF_data_list = new List<DataDetail>();
            List<TempData> BFS_data = new List<TempData>();
            foreach (var file in NoiseBeforeUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] BUF_data_array;
                BUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index, column_index);
                tmp.Data = BUF_data_array;
                tmp.SN = file.SN;
                BUF_data_list.Add(tmp);
                #region Prepare for SNR data
                TempData temp = new TempData();
                temp.objects = BUF_data_array;
                temp.SN = file.SN;
                BFS_data.Add(temp);
                #endregion
            }
            #endregion

            #region Get AfterUnderfill 
            List<DataDetail> AUF_data_list = new List<DataDetail>();
            List<TempData> AFS_data = new List<TempData>();
            foreach (var file in NoiseAfterUnderFill_FileList)
            {
                DataDetail tmp = new DataDetail();
                object[] AUF_data_array;
                AUF_data_array = GetSpecifyDataArray((object[,,])file.Data, sheet_index, column_index);
                tmp.Data = AUF_data_array;
                tmp.SN = file.SN;
                AUF_data_list.Add(tmp);
                #region Prepare for SNR data
                TempData temp = new TempData();
                temp.objects = AUF_data_array;
                temp.SN = file.SN;
                AFS_data.Add(temp);
                #endregion
            }
            #endregion

            #region Calculation
            var dictBUF = BUF_data_list.ToDictionary(d => d.SN);
            var dictAUF = AUF_data_list.ToDictionary(d => d.SN);
            List<DataDetail> diffList = new List<DataDetail>();

            foreach (var sn in dictBUF.Keys)
            {
                if (dictAUF.ContainsKey(sn))
                {
                    var dataBUF = dictBUF[sn].Data;
                    var dataAUF = dictAUF[sn].Data;
                    int len = Math.Min(dataBUF.Length, dataAUF.Length);
                    object[] diff = new object[len];
                    diff[0] = dataBUF[0];

                    for (int i = 1; i < len; i++)
                    {
                        double valBUF = dataBUF[i] != null ? Convert.ToDouble(dataBUF[i]) : 0;
                        double valAUF = dataAUF[i] != null ? Convert.ToDouble(dataAUF[i]) : 0;
                        diff[i] = valBUF - valAUF;
                    }

                    diffList.Add(new DataDetail { SN = sn, Data = diff });
                }
            }
            #endregion
            #region Prepare for SNR calculation
            Noise_list = new List<SNRData>();
            Noise_list = (from a in BFS_data
                          join b in AFS_data on a.SN equals b.SN into gj
                          from subB in gj.DefaultIfEmpty() // 沒對應就 subB 為 null
                          select new SNRData
                          {
                              SN = a.SN,
                              BeforeData = a.objects,
                              AfterData = subB?.objects ?? Array.Empty<object>()  // 找不到就 null
                          }).ToList();
            #endregion
            #region Merge by SN
            // X 軸欄位 = BUF.Count + 2空欄 + AUF.Count + 2空欄 + Diff.Count
            int colCount = BUF_data_list.Count + 2 + AUF_data_list.Count + 2 + diffList.Count;

            // Y 軸 = 資料最大長度
            int rowCount = Math.Max(
                BUF_data_list.Count > 0 ? BUF_data_list.Max(d => d.Data.Length) : 0,
                Math.Max(
                    AUF_data_list.Count > 0 ? AUF_data_list.Max(d => d.Data.Length) : 0,
                    diffList.Count > 0 ? diffList.Max(d => d.Data.Length) : 0
                )
            );


            object[,] result_array = new object[rowCount, colCount];

            // 塞 BUF (第一段欄)
            for (int c = 0; c < BUF_data_list.Count; c++)
            {
                var data = BUF_data_list[c].Data;
                for (int r = 0; r < data.Length; r++)
                    result_array[r, c] = data[r];
            }

            // 塞 AUF (空 2 欄後)
            int aufStart = BUF_data_list.Count + 2;
            for (int c = 0; c < AUF_data_list.Count; c++)
            {
                var data = AUF_data_list[c].Data;
                for (int r = 0; r < data.Length; r++)
                    result_array[r, aufStart + c] = data[r];
            }

            // 塞 Diff (空 2 + BUF長 + 2 + AUF長)
            int diffStart = BUF_data_list.Count + 2 + AUF_data_list.Count + 2;
            for (int c = 0; c < diffList.Count; c++)
            {
                var data = diffList[c].Data;
                for (int r = 0; r < data.Length; r++)
                {
                    if (data[r] is string)
                    {
                        result_array[r, diffStart + c] = data[r]; // 保留原字串
                    }
                    else
                    {
                        double val = Convert.ToDouble(data[r]);
                        result_array[r, diffStart + c] = (object)val;
                    }
                }
            }
            #endregion
            Tab2_array = result_array;

        }
        private void SNR_Table(out object[,] Tab_SNR_Array, List<SNRData> Signal_List, List<SNRData> Noise_List)
        {
            #region Get SNR before&after list
            List<object[]> signal_before_list = new List<object[]>();
            List<object[]> signal_after_list = new List<object[]>();
            List<object[]> noise_before_list = new List<object[]>();
            List<object[]> noise_after_list = new List<object[]>();
            Tab_SNR_Array = new object[Signal_List.Count, Noise_List.Count];
            foreach (var signal_temp in Signal_List)
            {
                object[] signal_before = signal_temp.BeforeData;
                object[] signal_after = signal_temp.AfterData;
                signal_before_list.Add(signal_before);
                signal_after_list.Add(signal_after);
            }
            foreach (var noise_temp in Noise_List)
            {
                object[] noise_before = noise_temp.BeforeData;
                object[] noise_after = noise_temp.AfterData;
                noise_before_list.Add(noise_before);
                noise_after_list.Add(noise_after);
            }
            List<object[]> snr_before_list = CalcSNR(signal_before_list, noise_before_list);
            List<object[]> snr_after_list = CalcSNR(signal_after_list, noise_after_list);
            #endregion
            #region calcu snr after&before diff
            List<object[]> snr_diff_list = CalcDiffBySN(snr_before_list, snr_after_list);
            #endregion
            #region merge
            int colCount = snr_before_list.Count + 2 + snr_after_list.Count + 2 + snr_diff_list.Count;
            int rowCount = Math.Max(
                snr_before_list.Count > 0 ? snr_before_list.Max(d => d.Length) : 0,
                Math.Max(
                    snr_after_list.Count > 0 ? snr_after_list.Max(d => d.Length) : 0,
                    snr_diff_list.Count > 0 ? snr_diff_list.Max(d => d.Length) : 0
                )
            );

            Tab_SNR_Array = new object[rowCount, colCount];

            // === Step 3: 塞 Before (BUF區) ===
            for (int c = 0; c < snr_before_list.Count; c++)
            {
                var data = snr_before_list[c];
                for (int r = 0; r < data.Length; r++)
                    Tab_SNR_Array[r, c] = data[r];
            }

            // === Step 4: 塞 After (AUF區) ===
            int afterStart = snr_before_list.Count + 2;
            for (int c = 0; c < snr_after_list.Count; c++)
            {
                var data = snr_after_list[c];
                for (int r = 0; r < data.Length; r++)
                    Tab_SNR_Array[r, afterStart + c] = data[r];
            }

            // === Step 5: 塞 Diff (Diff區) ===
            int diffStart = snr_before_list.Count + 2 + snr_after_list.Count + 2;
            for (int c = 0; c < snr_diff_list.Count; c++)
            {
                var data = snr_diff_list[c];
                for (int r = 0; r < data.Length; r++)
                {
                    if (data[r] is string)
                    {
                        Tab_SNR_Array[r, diffStart + c] = data[r]; // 保留 SN
                    }
                    else
                    {
                        double val = Convert.ToDouble(data[r]);
                        Tab_SNR_Array[r, diffStart + c] = (val == 0) ? null : (object)val;
                    }
                }
            }
            #endregion
        }
        private List<object[]> CalcSNR(List<object[]> signal_list, List<object[]> noise_list)
        {
            if (signal_list.Count != noise_list.Count)
                throw new ArgumentException("Signal 與 Noise 筆數不一致");

            List<object[]> snr_list = new List<object[]>();

            for (int i = 0; i < signal_list.Count; i++)
            {
                object[] sig_temp = signal_list[i];
                object[] noise_temp = noise_list[i];
                if (sig_temp.Length == 0 || noise_temp.Length == 0) { continue; }
                double noise = Convert.ToDouble(noise_temp[1]); // 固定取 index=1

                object[] snr_row = new object[sig_temp.Length];

                for (int j = 0; j < sig_temp.Length; j++)
                {
                    if (j == 0)
                    {
                        string name = sig_temp[j].ToString();
                        string[] x = name.Split("_");
                        x[1] = "SNR_" + x[1];
                        snr_row[j] = x[1]; 
                    }
                    else
                    {
                        double sig = Convert.ToDouble(sig_temp[j]);
                        snr_row[j] = sig - noise;
                    }
                }

                snr_list.Add(snr_row);
            }

            return snr_list;
        }
        public static List<object[]> CalcDiffBySN(List<object[]> beforeList, List<object[]> afterList)
        {
            List<object[]> diffList = new List<object[]>();

            // 建立 After 的查表 (SN -> object[])
            var afterDict = afterList
                .Where(a => a.Length > 0 && a[0] != null)
                .ToDictionary(a => a[0].ToString(), a => a);

            foreach (var before in beforeList)
            {
                if (before.Length == 0 || before[0] == null)
                    continue;

                string sn = before[0].ToString();

                // 找不到對應 SN 就跳過
                if (!afterDict.TryGetValue(sn, out var after))
                    continue;

                // 比對長度，以最小長度為基準
                int len = Math.Min(before.Length, after.Length);
                object[] diff = new object[len];
                diff[0] = sn; // 第一個是 SN，不計算差值

                for (int i = 1; i < len; i++)
                {
                    double b = Convert.ToDouble(before[i]);
                    double a = Convert.ToDouble(after[i]);
                    diff[i] = b - a; // Before 減 After
                }

                diffList.Add(diff);
            }

            return diffList;
        }
        private bool CreateExcelTable(object[,] data1, object[,] data2, object[,] data3)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    sfd.Title = "選擇輸出路徑";
                    sfd.FileName = "Summary.xlsx";

                    if (sfd.ShowDialog() != DialogResult.OK)
                        return false;

                    using (var package = new ExcelPackage())
                    {
                        var ws1 = package.Workbook.Worksheets.Add("Signal");
                        ws1.Cells[1, 1].LoadFromArrays(ToJaggedArray(data1));

                        var ws2 = package.Workbook.Worksheets.Add("Noise");
                        ws2.Cells[1, 1].LoadFromArrays(ToJaggedArray(data2));
                        var ws3 = package.Workbook.Worksheets.Add("SNR");
                        ws3.Cells[1, 1].LoadFromArrays(ToJaggedArray(data3));
                        package.SaveAs(new FileInfo(sfd.FileName));
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        private static object[][] ToJaggedArray(object[,] source)
        {
            int rows = source.GetLength(0);
            int cols = source.GetLength(1);
            var result = new object[rows][];
            for (int i = 0; i < rows; i++)
            {
                result[i] = new object[cols];
                for (int j = 0; j < cols; j++)
                    result[i][j] = source[i, j];
            }
            return result;
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
                    fileDetail.SN = SelectFileName_SN(files[i]);
                    file_list.Add(fileDetail);
                }
                Console.WriteLine($"TextBox:{uITextBox.Name},Path:{uITextBox.Text},ReadFileListData Done");
                switch (uITextBox.Name)
                {
                    case "uiTextBox3":
                        NoiseBeforeUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox4":
                        NoiseAfterUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox1":
                        SiganalBeforeUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox2":
                        SiganalAfterUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox5":
                        RangeProBeforeUnderFill_FileList = file_list;
                        break;
                    case "uiTextBox6":
                        RangeProAfterUnderFill_FileList = file_list;
                        break;
                }
            }
        }
        private void Start_Bt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Tab1_Sheet_tb.Text) || string.IsNullOrEmpty(Tab1_Column_tb.Text))
            {
                UIMessageBox.Show("tabpage1 參數未填寫");
                return;
            }
            if (string.IsNullOrEmpty(Tab2_Sheet_tb.Text) || string.IsNullOrEmpty(Tab2_Column_tb.Text))
            {
                UIMessageBox.Show("tabpage2 參數未填寫");
                return;
            }
            //Tab1_Result_Table(out object[,] tab1_array, out List<SNRData> Signal_Data_List);
            //Tab2_Result_Table(out object[,] tab2_array, out List<SNRData> Noise_Data_List);
            //SNR_Table(out object[,] snr_array, Signal_Data_List, Noise_Data_List);
            //if (CreateExcelTable(tab1_array, tab2_array, snr_array))
            //{
            //    UIMessageBox.Show("資料輸出完成!!");
            //}
            //else
            //{
            //    UIMessageBox.Show("資料輸出失敗!!");
            //}
            GetTableData(out List<Table> table_list);
            table_list = GetSNR(table_list);
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
            return sum - 1;
        }
        #endregion
        #region 小工具
        private string SelectFileName_SN(string FileName)
        {
            string name = Path.GetFileNameWithoutExtension(FileName);

            // 如果檔名有三段以上，用 '_' 分割
            var parts = name.Split('_');
            if (parts.Length >= 2)
            {
                // SN + 測項 = 前兩段組合
                return parts[0];
            }
            else
            {
                // 兩段直接回傳
                return name;
            }
        }
        #endregion
    }
}
