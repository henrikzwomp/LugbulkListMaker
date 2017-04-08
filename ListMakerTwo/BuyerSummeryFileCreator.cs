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

            var previous_color = "";

            int reservation_start_line = 14;

            int line_count = reservation_start_line;
            int column_offset = 0;
            foreach (var reservation in reservations
                .OrderBy(x => x.Element.BricklinkColor)
                .ThenBy(x => x.Element.ElementID)) // ToDo Test
            {
                var current_color = reservation.Element.BricklinkColor;

                if(current_color != previous_color)  // ToDo Test
                {
                    work_sheet.Range(line_count, 1 + column_offset, line_count, 3 + column_offset).Merge();
                    work_sheet.Cell(line_count, 1 + column_offset).Value = current_color;
                    work_sheet.Cell(line_count, 1 + column_offset).Style.Font.Bold = true;
                    work_sheet.Cell(line_count, 1 + column_offset).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                    line_count++;
                }

                work_sheet.Cell(line_count, 1 + column_offset).Value = reservation.Element.ElementID;  // ToDo Test
                work_sheet.Cell(line_count, 1 + column_offset).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

                work_sheet.Cell(line_count, 2 + column_offset).Value = reservation.Amount;
                work_sheet.Cell(line_count, 2 + column_offset).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                work_sheet.Cell(line_count, 1 + column_offset).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                work_sheet.Cell(line_count, 2 + column_offset).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                work_sheet.Cell(line_count, 3 + column_offset).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

                work_sheet.Cell(line_count, 1 + column_offset).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
                work_sheet.Cell(line_count, 2 + column_offset).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
                work_sheet.Cell(line_count, 3 + column_offset).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

                work_sheet.Cell(line_count, 1 + column_offset).Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
                work_sheet.Cell(line_count, 2 + column_offset).Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
                work_sheet.Cell(line_count, 3 + column_offset).Style.Border.SetRightBorder(XLBorderStyleValues.Thin);

                if(line_count == reservation_start_line) // ToDo Test
                {
                    work_sheet.Cell(line_count, 1 + column_offset).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
                    work_sheet.Cell(line_count, 2 + column_offset).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
                    work_sheet.Cell(line_count, 3 + column_offset).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
                }

                previous_color = current_color;
                line_count++;

                if(line_count >= 48) // ToDo Test
                {
                    line_count = reservation_start_line;
                    column_offset = 4;
                }
            }
        }
    }
}
