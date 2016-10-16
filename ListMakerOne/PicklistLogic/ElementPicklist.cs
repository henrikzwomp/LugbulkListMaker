using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerOne
{
    public class ElementPicklist : Element
    {
        public ElementPicklist()
        {
            Reservations = new List<ElementReservation>();
        }

        public IList<ElementReservation> Reservations { get; set; }
    }
}
