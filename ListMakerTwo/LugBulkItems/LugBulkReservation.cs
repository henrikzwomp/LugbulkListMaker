using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public struct LugBulkReservation
    {
        public LugBulkElement Element { get; set; }
        public LugBulkBuyer Buyer { get; set; }
        public int Amount { get; set; }
    }
}
