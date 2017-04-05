using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public class LugBulkPicklist : LugBulkElement
    {
        public LugBulkPicklist()
        {
            Reservations = new List<LugBulkReservation>();
        }

        public IList<LugBulkReservation> Reservations { get; set; }
    }
}
