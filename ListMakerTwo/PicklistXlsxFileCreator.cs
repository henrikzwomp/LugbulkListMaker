using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

namespace ListMakerTwo
{
    public class PicklistXlsxFileCreator // ToDo Test
    {
        public static void Create(IXLWorksheet work_sheet, LugBulkPicklist picklist) // string template_file_path
        {
            int line_count = 0;

            // Page one
            int page_one_buyers_row_start = 2;
            int page_one_buyers_row_end = 56;

            work_sheet.Cell(1, "B").Value = picklist.ElementID;
            work_sheet.Cell(2, "B").Value = picklist.BricklinkDescription;
            work_sheet.Cell(5, "B").Value = picklist.BricklinkColor;
            work_sheet.Cell(6, "B").Value = picklist.MaterialColor;

            for (int i = page_one_buyers_row_start; i <= page_one_buyers_row_end; i++)
            {
                if (line_count < picklist.Reservations.Count)
                {
                    work_sheet.Cell(i, "E").Value = picklist.Reservations[line_count].Receiver.Id;
                    work_sheet.Cell(i, "F").Value = picklist.Reservations[line_count].Receiver.Name;
                    work_sheet.Cell(i, "G").Value = picklist.Reservations[line_count].Amount;
                }
                else
                {
                    work_sheet.Cell(i, "H").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                    work_sheet.Cell(i, "G").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                    work_sheet.Cell(i, "F").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                    work_sheet.Cell(i, "E").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                }

                line_count++;
            }

            // Page two
            // ToDo: Refactor?
            int page_two_start_after = 56;
            int page_two_buyers_row_start = 58;
            int page_two_buyers_row_end = 112;

            for (int i = page_two_buyers_row_start; i <= page_two_buyers_row_end; i++)
            {
                if (line_count < picklist.Reservations.Count)
                {
                    work_sheet.Cell(i, "E").Value = picklist.Reservations[line_count].Receiver.Id;
                    work_sheet.Cell(i, "F").Value = picklist.Reservations[line_count].Receiver.Name;
                    work_sheet.Cell(i, "G").Value = picklist.Reservations[line_count].Amount;
                }
                else
                {
                    work_sheet.Cell(i, "H").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                    work_sheet.Cell(i, "G").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                    work_sheet.Cell(i, "F").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                    work_sheet.Cell(i, "E").Delete(XLShiftDeletedCells.ShiftCellsLeft);
                }

                line_count++;
            }

            work_sheet.Cell(1 + page_two_start_after, "B").Value = picklist.ElementID;
            work_sheet.Cell(2 + page_two_start_after, "B").Value = picklist.BricklinkDescription;
            work_sheet.Cell(5 + page_two_start_after, "B").Value = picklist.BricklinkColor;
            work_sheet.Cell(6 + page_two_start_after, "B").Value = picklist.MaterialColor;
        }
    }
}
