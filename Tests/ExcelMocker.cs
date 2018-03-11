using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Moq;
using ClosedXML.Excel;

namespace Tests
{
    class ExcelMocker
    {
        public static Mock<IXLRange> CreateMockRange(int column_start, int row_start, int column_end, int row_end)
        {
            var range = new Mock<IXLRange>();

            var range_column_start = new Mock<IXLRangeColumn>();
            range_column_start.Setup(x => x.ColumnNumber()).Returns(column_start);
            range.Setup(x => x.FirstColumn(null)).Returns(range_column_start.Object);

            var range_column_end = new Mock<IXLRangeColumn>();
            range_column_end.Setup(x => x.ColumnNumber()).Returns(column_end);
            range.Setup(x => x.LastColumn(null)).Returns(range_column_end.Object);

            var range_row_start = new Mock<IXLRangeRow>();
            range_row_start.Setup(x => x.RowNumber()).Returns(row_start);
            range.Setup(x => x.FirstRow(null)).Returns(range_row_start.Object);

            var range_row_end = new Mock<IXLRangeRow>();
            range_row_end.Setup(x => x.RowNumber()).Returns(row_end);
            range.Setup(x => x.LastRow(null)).Returns(range_row_end.Object);

            return range;
        }

        public static void CreateMockCell(string value, int row, int column, Mock<IXLWorksheet> sheet)
        {
            var cell = CreateMockCell(value, row, column);
            sheet.Setup(x => x.Cell(row, column)).Returns(cell.Object);
        }

        public static Mock<IXLCell> CreateMockCell(string value, int row, int column)
        {
            var cell = new Mock<IXLCell>();
            cell.Setup(x => x.Value).Returns(value);

            var address = new Mock<IXLAddress>();
            address.Setup(x => x.ColumnNumber).Returns(column);
            address.Setup(x => x.RowNumber).Returns(row);

            cell.Setup(x => x.Address).Returns(address.Object);


            return cell;
        }
    }
}
