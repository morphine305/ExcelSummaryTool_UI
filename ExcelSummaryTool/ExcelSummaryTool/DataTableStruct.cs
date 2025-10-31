using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSummaryTool
{
    public class FileDetail
    {
        public string FilePath { get; set; }
        public string FIleName { get; set; }
        public string SN { get; set; }
        public object Data { get; set; }
    }
    public class DataDetail
    {
        public string SN { get; set; }
        public object[] Data { get; set; }
    }
    public class DataDetail_Range
    {
        public string SN { get; set; }
        public object[,] Data { get; set; }
    }
    public class SNRData
    {
        public string SN {  set; get; }
        public object[] BeforeData { set; get; }
        public object[] AfterData { set; get; }
    }
    public class TempData
    {
        public string SN { set; get; }  
        public object[] objects { get; set; }
    }
    public class Table
    {
        public string Name { get; set; }
        public object[] Signal_Before { set; get; }
        public object[] Signal_After { set; get; }
        public object[] Noise_Before { set; get; }
        public object[] Noise_After { set; get; }
        public object[] SNR_Before { set; get; }
        public object[] SNR_After { set; get; }
    }
    public class Range_Table
    {
        public string Name { get; set; }
        public object[] AvgNoise_Array_Before { set; get; }
        public object[] AvgNoise_Array_After { set; get; }
        public object AvgNoise_Before { set; get; }
        public object AvgNoise_After { set; get; }

    }
}
