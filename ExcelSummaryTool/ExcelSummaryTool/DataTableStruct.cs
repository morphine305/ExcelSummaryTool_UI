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
}
