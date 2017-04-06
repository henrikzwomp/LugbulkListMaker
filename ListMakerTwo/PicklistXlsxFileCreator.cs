using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

namespace ListMakerTwo
{
    public class PicklistXlsxFileCreator
    {
        public static void Create(IXLWorksheet work_sheet, LugBulkPicklist picklist)
        {
            int line_count = 0;

            // Page one
            int page_one_buyers_row_start = 2;
            int page_one_buyers_row_end = 56;
            SetElementInfo(work_sheet, picklist, 1);

            for (int i = page_one_buyers_row_start; i <= page_one_buyers_row_end; i++)
            {
                SetReservationLine(work_sheet, picklist, line_count, i);
                line_count++;
            }

            // Page two
            int page_two_starts_at = 57;
            int page_two_ends_at = 112;
            int page_two_buyers_row_start = 58;
            int page_two_buyers_row_end = 112;
            int page_one_instruction_box_ends_at_row = 39;

            if (line_count >= picklist.Reservations.Count)
            {
                DeletePageTwoElementInfo(work_sheet, 
                    page_two_starts_at, page_two_ends_at,
                    page_one_instruction_box_ends_at_row);
            }
            else
            {
                SetElementInfo(work_sheet, picklist, page_two_starts_at);
            }

            for (int i = page_two_buyers_row_start; i <= page_two_buyers_row_end; i++)
            {
                SetReservationLine(work_sheet, picklist, line_count, i);
                line_count++;
            }
        }

        private static void DeletePageTwoElementInfo(IXLWorksheet work_sheet, 
            int page_two_start_at, int page_two_ends_at, int page_one_instruction_box_ends_at_row)
        {
            work_sheet.Rows(page_two_start_at, page_two_ends_at).Delete();

            while(work_sheet.MergedRanges.Any(x => x.FirstCell().Address.RowNumber > page_one_instruction_box_ends_at_row))
            {
                work_sheet.MergedRanges.Where(x => x.FirstCell().Address.RowNumber > page_one_instruction_box_ends_at_row).First()
                    .Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }
        }

        private static void SetElementInfo(IXLWorksheet work_sheet, LugBulkPicklist picklist, int origin_row)
        {
            work_sheet.Cell(origin_row + 0, "B").Value = picklist.ElementID;
            work_sheet.Cell(origin_row + 1, "B").Value = picklist.BricklinkDescription;
            work_sheet.Cell(origin_row + 4, "B").Value = picklist.BricklinkColor;
            work_sheet.Cell(origin_row + 5, "B").Value = picklist.MaterialColor;
        }

        private static void SetReservationLine(IXLWorksheet work_sheet, LugBulkPicklist picklist, 
            int line_count, int current_line)
        {
            if (line_count < picklist.Reservations.Count)
            {
                work_sheet.Cell(current_line, "E").Value = picklist.Reservations[line_count].Receiver.Id;
                work_sheet.Cell(current_line, "F").Value = picklist.Reservations[line_count].Receiver.Name;
                work_sheet.Cell(current_line, "G").Value = picklist.Reservations[line_count].Amount;
            }
            else
            {
                work_sheet.Cell(current_line, "H").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                work_sheet.Cell(current_line, "G").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                work_sheet.Cell(current_line, "F").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                work_sheet.Cell(current_line, "E").Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }
        }
    }
}
