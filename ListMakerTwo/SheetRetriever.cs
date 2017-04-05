using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    public class SheetRetriever
    {
        public static IXLWorksheet Get(string source_file_path, string work_sheet_name)
        {
            var workbook = new XLWorkbook(source_file_path);
            return workbook.Worksheets.First(x => x.Name == work_sheet_name);
        }
    }
}
