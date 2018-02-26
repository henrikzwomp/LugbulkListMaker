using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Moq;
using ClosedXML.Excel;
using ListMakerTwo;

namespace Tests.ListMakerTwo
{
    [TestFixture]
    public class SourceReaderHelperTests
    {
        private static Mock<IXLRange> CreateMockRange(int column_start, int row_start, int column_end, int row_end)
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

        private void CreateMockCell(string value, int row, int column, Mock<IXLWorksheet> sheet)
        {
            var cell = new Mock<IXLCell>(); cell.Setup(x => x.Value).Returns(value);
            sheet.Setup(x => x.Cell(row, column)).Returns(cell.Object);
        }

        [Test]
        public void CanGetGetBottomRightCrossRangePositions()
        {
            /*
                A   B   C   D   E
            1           .   .
            2
            3   .       x   x
            4   .       x   x
            5

            */

            var C1D1 = CreateMockRange(3, 1, 4, 1);
            var A3A4 = CreateMockRange(1, 3, 1, 4);

            var result1 = SourceReaderHelper.GetBottomRightCrossRangePositions(C1D1.Object, A3A4.Object);
            var result2 = SourceReaderHelper.GetBottomRightCrossRangePositions(A3A4.Object, C1D1.Object);

            Assert.That(result1.Count, Is.EqualTo(4));
            Assert.That(result1.Count(x => x.Row == 3 && x.Column == 3), Is.EqualTo(1));
            Assert.That(result1.Count(x => x.Row == 3 && x.Column == 4), Is.EqualTo(1));
            Assert.That(result1.Count(x => x.Row == 4 && x.Column == 3), Is.EqualTo(1));
            Assert.That(result1.Count(x => x.Row == 4 && x.Column == 4), Is.EqualTo(1));

            Assert.That(result2.Count, Is.EqualTo(4));
            Assert.That(result2.Count(x => x.Row == 3 && x.Column == 3), Is.EqualTo(1));
            Assert.That(result2.Count(x => x.Row == 3 && x.Column == 4), Is.EqualTo(1));
            Assert.That(result2.Count(x => x.Row == 4 && x.Column == 3), Is.EqualTo(1));
            Assert.That(result2.Count(x => x.Row == 4 && x.Column == 4), Is.EqualTo(1));
        }

        [Test]
        public void CanGetTitlePositionForValuePosition()
        {
            // GetTitleValueForReservation
            /*
                A   B   C   D   E
            1           .   .
            2
            3   .       C3  D3
            4   .       C4  D4
            5

            */

            var C1D1 = CreateMockRange(3, 1, 4, 1);
            var A3A4 = CreateMockRange(1, 3, 1, 4);

            var title_for_C = SourceReaderHelper.GetTitlePositionForValuePosition(
                new CellPosition() { Column = 3, Row = 3 }, C1D1.Object);

            var title_for_D = SourceReaderHelper.GetTitlePositionForValuePosition(
                new CellPosition() { Column = 4, Row = 4 }, C1D1.Object);

            var title_for_3 = SourceReaderHelper.GetTitlePositionForValuePosition(
                new CellPosition() { Column = 4, Row = 3 }, A3A4.Object);

            var title_for_4 = SourceReaderHelper.GetTitlePositionForValuePosition(
                new CellPosition() { Column = 3, Row = 4 }, A3A4.Object);

            Assert.That(title_for_C.Row, Is.EqualTo(1));
            Assert.That(title_for_C.Column, Is.EqualTo(3));
            Assert.That(title_for_D.Row, Is.EqualTo(1));
            Assert.That(title_for_D.Column, Is.EqualTo(4));
            Assert.That(title_for_3.Row, Is.EqualTo(3));
            Assert.That(title_for_3.Column, Is.EqualTo(1));
            Assert.That(title_for_4.Row, Is.EqualTo(4));
            Assert.That(title_for_4.Column, Is.EqualTo(1));
        }

        [Test]
        public void CanGetValuesForTitlePosition() 
        {
            /*
                A   B   C   D
            1       BB  CC
            2   22  B2  C2
            3   33  B3  C3
            4
            */

            var title_pos_1 = new CellPosition() { Row = 1, Column = 2 };
            var title_pos_2 = new CellPosition() { Row = 2, Column = 1 };

            var values_start_pos = new CellPosition() { Row = 2, Column = 2 };
            var values_end_pos = new CellPosition() { Row = 3, Column = 3 };

            var work_sheet = new Mock<IXLWorksheet>();

            CreateMockCell("B2", 2, 2, work_sheet);
            CreateMockCell("B3", 3, 2, work_sheet);
            CreateMockCell("C2", 2, 3, work_sheet);
            CreateMockCell("C3", 3, 3, work_sheet);

            var result1 = SourceReaderHelper.GetValuesForTitlePosition(title_pos_1,
                values_start_pos, values_end_pos, work_sheet.Object);

            var result2 = SourceReaderHelper.GetValuesForTitlePosition(title_pos_2,
                values_start_pos, values_end_pos, work_sheet.Object);

            Assert.That(result1.Count, Is.EqualTo(2));
            Assert.That(result1.Count(x => x == "B2"), Is.EqualTo(1));
            Assert.That(result1.Count(x => x == "B3"), Is.EqualTo(1));

            Assert.That(result2.Count, Is.EqualTo(2));
            Assert.That(result2.Count(x => x == "B2"), Is.EqualTo(1));
            Assert.That(result2.Count(x => x == "C2"), Is.EqualTo(1));


        }
    }
}
