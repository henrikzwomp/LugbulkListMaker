using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    public class InputParameters
    {
        // Required
        public string SourceFileName { get; set; }
        public string WorksheetName { get; set; }
        public IXLRange BuyersSpan { get; set; }
        public IXLRange ElementIdSpan { get; set; }
        public IXLRange BrickLinkDescriptionSpan { get; set; }
        public IXLRange BrickLinkIdSpan { get; set; }
        public IXLRange BrickLinkColorSpan { get; set; }
        public IXLRange TlgColorSpan { get; set; }
    }
}
