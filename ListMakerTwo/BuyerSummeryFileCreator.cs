using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    class BuyerSummeryFileCreator
    {
        internal static void Create(IXLWorksheet work_sheet, LugBulkBuyer buyer, 
            IEnumerable<LugBulkReservation> reservations) // ToDo Test
        {
            work_sheet.Cell(1, "A").Value = buyer.Id + " - " + buyer.Name;

            int line_count = 4;

            var last_color = "";

            var alt_color_1 = XLColor.White;
            var alt_color_2 = work_sheet.Cell(line_count, 1).Style.Fill.BackgroundColor;
            var current_color = alt_color_2;


            foreach (var reservation in reservations
                .OrderBy(x => x.Element.BricklinkColor)
                .ThenBy(x => x.Element.BricklinkDescription))
            {
                if (reservation.Element.BricklinkColor != last_color)
                {
                    if (current_color == alt_color_1)
                        current_color = alt_color_2;
                    else
                        current_color = alt_color_1;
                }

                SetCell(work_sheet.Cell(line_count, 1), reservation.Element.ElementID, current_color);
                SetCell(work_sheet.Cell(line_count, 2), reservation.Element.BricklinkColor, current_color);
                SetCell(work_sheet.Cell(line_count, 3), reservation.Element.BricklinkDescription, current_color);
                SetCell(work_sheet.Cell(line_count, 4), reservation.Amount.ToString(), current_color);
                line_count++;

                last_color = reservation.Element.BricklinkColor;
            }


        }

        private static void SetCell(IXLCell cell, string value, XLColor color)
        {
            cell.Value = value;

            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            cell.Style.Fill.BackgroundColor = color;
        }
        
    }
}
