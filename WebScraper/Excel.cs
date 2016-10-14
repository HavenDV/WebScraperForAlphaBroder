using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.Data;
using System.IO;
using System.Diagnostics;

namespace WebScraper
{
    static class Excel
    {
        public static DataTable LoadExcelFile(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var extension = Path.GetExtension(path);
                var isXls = string.Compare(extension, ".xls", StringComparison.OrdinalIgnoreCase) == 0;
                var isXlsx = string.Compare(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) == 0;
                if (!isXls && !isXlsx)
                {
                    //return LoadTextFile(path);
                }

                var reader = isXls ?
                    ExcelReaderFactory.CreateBinaryReader(stream) :
                    ExcelReaderFactory.CreateOpenXmlReader(stream);
                reader.IsFirstRowAsColumnNames = true;
                var data = reader.AsDataSet();
                Debug.Assert(data != null);
                Debug.Assert(data.Tables != null);
                Debug.Assert(data.Tables.Count > 0);
                return data.Tables[0];
            }
        }
    }
}
