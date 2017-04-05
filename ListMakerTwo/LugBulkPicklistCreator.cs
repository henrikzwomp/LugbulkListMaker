using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public class LugBulkPicklistCreator // ToDo Copy tests
    {
        public static IList<LugBulkPicklist> CreateLists(IList<LugBulkReservation> reservations, IList<LugBulkElement> elements)
        {
            var picklists = new SortedList<string, LugBulkPicklist>();

            foreach (var element_res in reservations)
            {
                if (!picklists.Keys.Any(x => x == element_res.ElementID))
                {
                    var new_picklist = new LugBulkPicklist() { ElementID = element_res.ElementID };

                    if (elements.Any(x => x.ElementID == element_res.ElementID))
                    {
                        var element = elements.Where(x => x.ElementID == element_res.ElementID).First();
                        new_picklist.BricklinkDescription = element.BricklinkDescription;
                        new_picklist.BricklinkColor = element.BricklinkColor;
                        new_picklist.MaterialColor = element.MaterialColor;
                    }

                    picklists.Add(element_res.ElementID, new_picklist);
                }

                var pick_list = picklists[element_res.ElementID];

                pick_list.Reservations.Add(element_res);
            }

            Parallel.For(0, picklists.Count,
                   index =>
                   {
                       picklists.Values[index].Reservations = new List<LugBulkReservation>(
                           picklists.Values[index].Reservations.OrderBy(x => x.Amount).ThenBy(x => x.Receiver.Name));
                   });

            return picklists.Values.OrderBy(x => x.ElementID).ToList<LugBulkPicklist>();
        }
    }
}
