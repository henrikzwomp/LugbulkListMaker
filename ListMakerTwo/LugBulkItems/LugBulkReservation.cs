using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public struct LugBulkReservation
    {
        public string ElementID { get; set; }
        public LugBulkReceiver Receiver { get; set; }
        public int Amount { get; set; }
    }
}
