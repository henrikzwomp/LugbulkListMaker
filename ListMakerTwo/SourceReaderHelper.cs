using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    public class SourceReaderHelper
    {
        public static CellPosition GetTitlePositionForValuePosition(CellPosition value_pos, IXLRange title_span)
        {
            var first_column = title_span.FirstColumn().ColumnNumber();
            var last_column = title_span.LastColumn().ColumnNumber();
            var first_row = title_span.FirstRow().RowNumber();
            var last_row = title_span.LastRow().RowNumber();

            if (first_column <= value_pos.Column && value_pos.Column <= last_column)
                return new CellPosition() { Row = first_row, Column = value_pos.Column };

            if (first_row <= value_pos.Row && value_pos.Row <= last_row)
                return new CellPosition() { Row = value_pos.Row, Column = first_column };

            throw new Exception("Failed GetTitlePositionForValuePosition");
        }

        public static IList<string> GetValuesForTitlePosition(CellPosition title_pos, 
            CellPosition values_start_pos, CellPosition values_end_pos, IXLWorksheet work_sheet)
        {
            var result = new List<string>();

            // If title Column is in span
            if (values_start_pos.Column <= title_pos.Column && title_pos.Column <= values_end_pos.Column)
            {
                for(int i = values_start_pos.Column; i <= values_end_pos.Column; i++)
                {
                    result.Add(work_sheet.Cell(i, title_pos.Column).Value.ToString().Trim());
                }
            }
            // If title Row is in span
            else if (values_start_pos.Row <= title_pos.Row && title_pos.Row <= values_end_pos.Row)
            {
                for (int i = values_start_pos.Row; i <= values_end_pos.Row; i++)
                {
                    result.Add(work_sheet.Cell(title_pos.Row, i).Value.ToString().Trim());
                }
            }

            return result;
        }

        public static IList<CellPosition> GetBottomRightCrossRangePositions(IXLRange range1, IXLRange range2)
        {
            CellPosition start_pos = null;
            CellPosition end_pos = null;

            GetBottomRightCrossRangeStartEndPositions(range1, range2, out start_pos, out end_pos);

            var result = new List<CellPosition>();

            for(int x = start_pos.Column; x <= end_pos.Column; x++)
            {
                for(int y = start_pos.Row; y <= end_pos.Row; y++)
                {
                    result.Add(new CellPosition() { Column = x, Row = y });
                }
            }

            return result;
        }

        public static void GetBottomRightCrossRangeStartEndPositions(IXLRange range1, IXLRange range2
            , out CellPosition start_pos, out CellPosition end_pos)
        {
            start_pos = null;
            end_pos = null;

            var first_column_1 = range1.FirstColumn().ColumnNumber();
            var last_column_1 = range1.LastColumn().ColumnNumber();
            var first_row_1 = range1.FirstRow().RowNumber();
            var last_row_1 = range1.LastRow().RowNumber();

            var first_column_2 = range2.FirstColumn().ColumnNumber();
            var last_column_2 = range2.LastColumn().ColumnNumber();
            var first_row_2 = range2.FirstRow().RowNumber();
            var last_row_2 = range2.LastRow().RowNumber();

            var start_pos_candidate_1 = new CellPosition() { Row = first_row_1, Column = first_column_2 };
            var start_pos_candidate_2 = new CellPosition() { Row = first_column_1, Column = first_row_2 };

            var end_pos_candidate_1 = new CellPosition() { Row = last_row_1, Column = last_column_2 };
            var end_pos_candidate_2 = new CellPosition() { Row = last_column_1, Column = last_row_2 };

            // Not above or to the left of a span.
            if (start_pos_candidate_1.Column >= first_column_1 && start_pos_candidate_1.Row >= first_row_1)
                start_pos = start_pos_candidate_1;
            else
                start_pos = start_pos_candidate_2;

            if (end_pos_candidate_1.Column >= last_column_1 && end_pos_candidate_1.Row >= last_row_1)
                end_pos = end_pos_candidate_1;
            else
                end_pos = end_pos_candidate_2;
        }
    }
}
