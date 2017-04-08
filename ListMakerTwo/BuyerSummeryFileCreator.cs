using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    class BuyerSummeryFileCreator // ToDo: New template must be designed. 
    {
        internal static void Create(IXLWorksheet work_sheet, LugBulkBuyer buyer, 
            IEnumerable<LugBulkReservation> reservations) // ToDo Test
        {
            work_sheet.Cell(1, "A").Value = buyer.Id; // ToDo Test
            work_sheet.Cell(11, "A").Value = buyer.Name; // ToDo Test

            int line_count = 15;
            foreach (var reservation in reservations
                .OrderBy(x => x.Element.BricklinkColor)
                .ThenBy(x => x.Element.ElementID)) // ToDo Test
            {
                work_sheet.Cell(line_count, "A").Value = reservation.Element.BricklinkColor;
                work_sheet.Cell(line_count, "B").Value = reservation.Element.ElementID;
                work_sheet.Cell(line_count, "C").Value = reservation.Amount;
                line_count++;
            }

            if (line_count >= 50) // ToDo Test?
                return;

            for(int i = 50; i >= line_count; i--) // ToDo Test
            {
                work_sheet.Row(i).Delete();
            }
        }
    }
}
