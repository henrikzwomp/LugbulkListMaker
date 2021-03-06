﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public class LugBulkPicklistCreator
    {
        public static IList<LugBulkPicklist> CreateLists(IList<LugBulkReservation> reservations, IList<LugBulkElement> elements)
        {
            var picklists = new SortedList<string, LugBulkPicklist>();

            foreach (var element_res in reservations)
            {
                if (!picklists.Keys.Any(x => x == element_res.Element.ElementID))
                {
                    var new_picklist = new LugBulkPicklist() { ElementID = element_res.Element.ElementID };

                    if (elements.Any(x => x.ElementID == element_res.Element.ElementID))
                    {
                        var element = elements.Where(x => x.ElementID == element_res.Element.ElementID).First();
                        new_picklist.BricklinkDescription = element.BricklinkDescription;
                        new_picklist.BricklinkColor = element.BricklinkColor;
                        new_picklist.MaterialColor = element.MaterialColor;
                    }

                    picklists.Add(element_res.Element.ElementID, new_picklist);
                }

                var pick_list = picklists[element_res.Element.ElementID];

                pick_list.Reservations.Add(element_res);
            }

            Parallel.For(0, picklists.Count,
                   index =>
                   {
                       picklists.Values[index].Reservations = new List<LugBulkReservation>(
                           picklists.Values[index].Reservations.OrderBy(x => x.Amount).ThenBy(x => x.Buyer.Name));
                   });

            return picklists.Values.OrderBy(x => x.ElementID).ToList<LugBulkPicklist>();
        }
    }
}
