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
}
