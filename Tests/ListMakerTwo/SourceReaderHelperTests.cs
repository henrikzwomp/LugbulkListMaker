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
        [Test]
        public void GetCrossRangePositions_CanGetGetBottomRightCrossRangePositions()
        {
            /*
                A   B   C   D   E
            1           .   .
            2
            3   .       x   x
            4   .       x   x
            5

            */

            var C1D1 = ExcelMocker.CreateMockRange(3, 1, 4, 1);
            var A3A4 = ExcelMocker.CreateMockRange(1, 3, 1, 4);

            var result1 = SourceReaderHelper.GetCrossRangePositions(C1D1.Object, A3A4.Object);
            var result2 = SourceReaderHelper.GetCrossRangePositions(A3A4.Object, C1D1.Object);

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
        public void GetCrossRangePositions_CanGetGetTopRightCrossRangePositions()
        {
            /*
                A   B   C   D   E
            1           
            2
            3   .       x   x
            4   .       x   x
            5
            6           .   .

            */

            var C6D6 = ExcelMocker.CreateMockRange(3, 6, 4, 6);
            var A3A4 = ExcelMocker.CreateMockRange(1, 3, 1, 4);

            var result1 = SourceReaderHelper.GetCrossRangePositions(C6D6.Object, A3A4.Object);
            var result2 = SourceReaderHelper.GetCrossRangePositions(A3A4.Object, C6D6.Object);

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
        public void CanGetBiggestCrossRangePositions_RealExample2017()
        {
            var _fake_sheet = new XLWorkbook().AddWorksheet("Fake");
            var BuyersSpan = _fake_sheet.Range("K87:FW87");
            var ElementIdSpan = _fake_sheet.Range("D2:D86");

            // K2 FW86

            var result1 = SourceReaderHelper.GetCrossRangePositions(BuyersSpan, ElementIdSpan);
            var result2 = SourceReaderHelper.GetCrossRangePositions(ElementIdSpan, BuyersSpan);

            Assert.That(result1.Count, Is.EqualTo(14365));
            Assert.That(result1.Count(x => x.Row == 2 && x.Column == 11), Is.EqualTo(1));
            Assert.That(result1.Count(x => x.Row == 86 && x.Column == 179), Is.EqualTo(1));

            Assert.That(result2.Count, Is.EqualTo(14365));
            Assert.That(result2.Count(x => x.Row == 2 && x.Column == 11), Is.EqualTo(1));
            Assert.That(result2.Count(x => x.Row == 86 && x.Column == 179), Is.EqualTo(1));
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

            var C1D1 = ExcelMocker.CreateMockRange(3, 1, 4, 1);
            var A3A4 = ExcelMocker.CreateMockRange(1, 3, 1, 4);

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
                A   B   C   D   E
            1       BB  CC  DD
            2   22  B2  C2  D2
            3   33  B3  C3  D3
            4
            */

            var title_pos_1 = new CellPosition() { Row = 1, Column = 2 };
            var title_pos_2 = new CellPosition() { Row = 2, Column = 1 };

            var values_start_pos = new CellPosition() { Row = 2, Column = 2 };
            var values_end_pos = new CellPosition() { Row = 3, Column = 4 };

            var work_sheet = new Mock<IXLWorksheet>();

            ExcelMocker.CreateMockCell("B2", 2, 2, work_sheet);
            ExcelMocker.CreateMockCell("B3", 3, 2, work_sheet);
            ExcelMocker.CreateMockCell("C2", 2, 3, work_sheet);
            ExcelMocker.CreateMockCell("C3", 3, 3, work_sheet);
            ExcelMocker.CreateMockCell("D2", 2, 4, work_sheet);
            ExcelMocker.CreateMockCell("D3", 3, 4, work_sheet);

            var result1 = SourceReaderHelper.GetValuesForTitlePosition(title_pos_1,
                values_start_pos, values_end_pos, work_sheet.Object);

            var result2 = SourceReaderHelper.GetValuesForTitlePosition(title_pos_2,
                values_start_pos, values_end_pos, work_sheet.Object);

            Assert.That(result1.Count, Is.EqualTo(2));
            Assert.That(result1.Count(x => x == "B2"), Is.EqualTo(1));
            Assert.That(result1.Count(x => x == "B3"), Is.EqualTo(1));

            Assert.That(result2.Count, Is.EqualTo(3));
            Assert.That(result2.Count(x => x == "B2"), Is.EqualTo(1));
            Assert.That(result2.Count(x => x == "C2"), Is.EqualTo(1));
            Assert.That(result2.Count(x => x == "D2"), Is.EqualTo(1));
        }

        [Test]
        public void CanGetValuesForTitlePosition_Pivoted()
        {
            /*
                A   B   C   D   
            1       BB  CC  
            2   22  B2  C2  
            3   33  B3  C3  
            4   44  B4  C4
            5
            */

            var title_pos_1 = new CellPosition() { Row = 1, Column = 2 };
            var title_pos_2 = new CellPosition() { Row = 2, Column = 1 };

            var values_start_pos = new CellPosition() { Row = 2, Column = 2 };
            var values_end_pos = new CellPosition() { Row = 4, Column = 3 };

            var work_sheet = new Mock<IXLWorksheet>();

            ExcelMocker.CreateMockCell("B2", 2, 2, work_sheet);
            ExcelMocker.CreateMockCell("B3", 3, 2, work_sheet);
            ExcelMocker.CreateMockCell("B4", 4, 2, work_sheet);
            ExcelMocker.CreateMockCell("C2", 2, 3, work_sheet);
            ExcelMocker.CreateMockCell("C3", 3, 3, work_sheet);
            ExcelMocker.CreateMockCell("C4", 4, 3, work_sheet);

            var result1 = SourceReaderHelper.GetValuesForTitlePosition(title_pos_1,
                values_start_pos, values_end_pos, work_sheet.Object);

            var result2 = SourceReaderHelper.GetValuesForTitlePosition(title_pos_2,
                values_start_pos, values_end_pos, work_sheet.Object);

            Assert.That(result1.Count, Is.EqualTo(3));
            Assert.That(result1.Count(x => x == "B2"), Is.EqualTo(1));
            Assert.That(result1.Count(x => x == "B3"), Is.EqualTo(1));
            Assert.That(result1.Count(x => x == "B4"), Is.EqualTo(1));

            Assert.That(result2.Count, Is.EqualTo(2));
            Assert.That(result2.Count(x => x == "B2"), Is.EqualTo(1));
            Assert.That(result2.Count(x => x == "C2"), Is.EqualTo(1));
        }

        [Test]
        public void GetCrossRangePositions_WillFailIfSpansAreMoreThanOneCellWideBothWays()
        {
            var C1D1_Thin = ExcelMocker.CreateMockRange(3, 1, 4, 1);
            var A3A4_Thin = ExcelMocker.CreateMockRange(1, 3, 1, 4);
            var C1D2_Wide = ExcelMocker.CreateMockRange(3, 1, 4, 2);
            var A3B4_Wide = ExcelMocker.CreateMockRange(1, 3, 2, 4);

            Assert.Throws(typeof(Exception), () => { SourceReaderHelper.GetCrossRangePositions(C1D1_Thin.Object, A3B4_Wide.Object); });
            Assert.Throws(typeof(Exception), () => { SourceReaderHelper.GetCrossRangePositions(A3A4_Thin.Object, C1D2_Wide.Object); });
            Assert.Throws(typeof(Exception), () => { SourceReaderHelper.GetCrossRangePositions(C1D2_Wide.Object, C1D1_Thin.Object); });
            Assert.Throws(typeof(Exception), () => { SourceReaderHelper.GetCrossRangePositions(A3B4_Wide.Object, A3A4_Thin.Object); });
        }

        [Test]
        public void GetValuesForTitlePosition_RealExample2017()
        {
            var sheet = new XLWorkbook().AddWorksheet("Test");

            var reservation_values = SourceReaderHelper.GetValuesForTitlePosition(
                        new CellPosition() { Column = 11, Row = 87 },
                        new CellPosition() { Column = 11, Row = 2 },
                        new CellPosition() { Column = 179, Row = 86 },
                        sheet);

            Assert.That(reservation_values.Count, Is.EqualTo(85));
        }
    }
}
