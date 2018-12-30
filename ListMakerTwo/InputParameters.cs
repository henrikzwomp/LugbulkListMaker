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
        public string SourceFileName { get; set; }
        public string WorksheetName { get; set; }
        public string BuyersSpan { get; set; }
        public string ElementIdSpan { get; set; }
        public string BrickLinkDescriptionSpan { get; set; }
        public string BrickLinkIdSpan { get; set; }
        public string BrickLinkColorSpan { get; set; }
        public string TlgColorSpan { get; set; }
        public string OutputFolder { get; set; }
    }
}
