using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public class InputParameters
    {
        // Required
        public string SourceFileName { get; set; }
        public string WorksheetName { get; set; }
        public string ElementRowSpan { get; set; }
        public int BuyersRow { get; set; }
        public string BuyersColumnSpan { get; set; }
        public string ElementIdColumn { get; set; }

        // ToDo: Optional
        public string BrickLinkDescriptionColumn { get; set; }
        public string BrickLinkIdColumn { get; set; }
        public string BrickLinkColorColumn { get; set; }
        public string TlgColorColumn { get; set; } // ToDo: Not used 
    }
}
