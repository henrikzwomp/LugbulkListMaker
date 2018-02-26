using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Moq;
using ListMakerTwo;
using ClosedXML.Excel;

namespace Tests.ListMakerTwo
{
    [TestFixture]
    public class SourceReaderTests
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
        public void SourceReaderGen2_CanGetBuyers()
        {
            /*
                _       Henrik      Alice   Simpson
                111     1           0       1 
                222     1           1       0
                222     0           1       0
            */

            var range_A2A4 = CreateMockRange(1, 2, 1, 4);
            var range_B1D1 = CreateMockRange(2, 1, 4, 1);

            var parameters = new InputParameters();
            parameters.ElementIdSpan = range_A2A4.Object;
            parameters.BuyersSpan = range_B1D1.Object;

            var sheet = new Mock<IXLWorksheet>();

            CreateMockCell("Henrik", 1, 2, sheet);
            CreateMockCell("Alice", 1, 3, sheet);
            CreateMockCell("Simpson", 1, 4, sheet);

            CreateMockCell("1", 2, 2, sheet);
            CreateMockCell("0", 2, 3, sheet);
            CreateMockCell("1", 2, 4, sheet);

            CreateMockCell("1", 3, 2, sheet);
            CreateMockCell("1", 3, 3, sheet);
            CreateMockCell("0", 3, 4, sheet);

            CreateMockCell("0", 4, 2, sheet);
            CreateMockCell("1", 4, 3, sheet);
            CreateMockCell("0", 4, 4, sheet);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetBuyers();

            Assert.That(result.Count, Is.EqualTo(3));
            Assert.That(result[0].Name, Is.EqualTo("Henrik"));
            Assert.That(result[1].Name, Is.EqualTo("Alice"));
            Assert.That(result[2].Name, Is.EqualTo("Simpson"));
            Assert.That(result[0].Id, Is.EqualTo(100));
            Assert.That(result[1].Id, Is.EqualTo(101));
            Assert.That(result[2].Id, Is.EqualTo(102));
        }

        [Test]
        public void SourceReaderGen2_CanGetElements()
        {
            /*
            ElementID   BL Desc BL Id   BL Color    TLG Color    Henrik      Alice   Simpson
            111         Brick1  BB1     Red         Real Read    1           0       1 
            222         Brick2  BB2     Blue        Bright Blue  1           1       0
            333         Brick3  BB3     Green       Dark Green   0           1       0
           */

            var range_A2A4 = CreateMockRange(1,2,1,4);
            var range_B2B4 = CreateMockRange(2, 2, 2, 4);
            var range_C2C4 = CreateMockRange(3, 2, 3, 4);
            var range_D2D4 = CreateMockRange(4, 2, 4, 4);
            var range_E2E4 = CreateMockRange(5, 2, 5, 4);

            var parameters = new InputParameters();

            parameters.ElementIdSpan = range_A2A4.Object;
            parameters.BrickLinkDescriptionSpan = range_B2B4.Object;
            parameters.BrickLinkIdSpan = range_C2C4.Object;
            parameters.BrickLinkColorSpan = range_D2D4.Object;
            parameters.TlgColorSpan = range_E2E4.Object;

            var sheet = new Mock<IXLWorksheet>();
            CreateMockCell("111", 2, 1, sheet);
            CreateMockCell("222", 3, 1, sheet);
            CreateMockCell("333", 4, 1, sheet);

            CreateMockCell("Brick1", 2, 2, sheet);
            CreateMockCell("Brick2", 3, 2, sheet);
            CreateMockCell("Brick3", 4, 2, sheet);

            CreateMockCell("BB1", 2, 3, sheet);
            CreateMockCell("BB2", 3, 3, sheet);
            CreateMockCell("BB3", 4, 3, sheet);

            CreateMockCell("Red", 2, 4, sheet);
            CreateMockCell("Blue", 3, 4, sheet);
            CreateMockCell("Green", 4, 4, sheet);

            CreateMockCell("Real Red", 2, 5, sheet);
            CreateMockCell("Bright Blue", 3, 5, sheet);
            CreateMockCell("Dark Green", 4, 5, sheet);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetElements();

            Assert.That(result.Count, Is.EqualTo(3));

            Assert.That(result[0].ElementID, Is.EqualTo("111"));
            Assert.That(result[0].BricklinkDescription, Is.EqualTo("Brick1"));
            Assert.That(result[0].BricklinkId, Is.EqualTo("BB1"));
            Assert.That(result[0].BricklinkColor, Is.EqualTo("Red"));
            Assert.That(result[0].MaterialColor, Is.EqualTo("Real Red"));

            Assert.That(result[1].ElementID, Is.EqualTo("222"));
            Assert.That(result[1].BricklinkDescription, Is.EqualTo("Brick2"));
            Assert.That(result[1].BricklinkId, Is.EqualTo("BB2"));
            Assert.That(result[1].BricklinkColor, Is.EqualTo("Blue"));
            Assert.That(result[1].MaterialColor, Is.EqualTo("Bright Blue"));

            Assert.That(result[2].ElementID, Is.EqualTo("333"));
            Assert.That(result[2].BricklinkDescription, Is.EqualTo("Brick3"));
            Assert.That(result[2].BricklinkId, Is.EqualTo("BB3"));
            Assert.That(result[2].BricklinkColor, Is.EqualTo("Green"));
            Assert.That(result[2].MaterialColor, Is.EqualTo("Dark Green"));
        }

        // ToDo Same test but with pivited table
        [Test]
        public void SourceReaderGen2_CanGetBuyersWithAReservations()
        {
            /*
                _       Henrik      Alice   Simpson
                111     1           0       1 
                222     1           -       0
                333                         0
            */

            var range_A2A4 = CreateMockRange(1, 2, 1, 4);
            var range_B1D1 = CreateMockRange(2, 1, 4, 1);

            var parameters = new InputParameters();
            parameters.ElementIdSpan = range_A2A4.Object;
            parameters.BuyersSpan = range_B1D1.Object;
            
            // Reusing columns to simplify test
            parameters.BrickLinkDescriptionSpan = range_A2A4.Object;
            parameters.BrickLinkIdSpan = range_A2A4.Object;
            parameters.BrickLinkColorSpan = range_A2A4.Object;
            parameters.TlgColorSpan = range_A2A4.Object;

            var sheet = new Mock<IXLWorksheet>();
            
            CreateMockCell("Henrik", 1, 2, sheet);
            CreateMockCell("Alice", 1, 3, sheet);
            CreateMockCell("Simpson", 1, 4, sheet);

            CreateMockCell("111", 2, 1, sheet);
            CreateMockCell("222", 3, 1, sheet);
            CreateMockCell("333", 4, 1, sheet);

            CreateMockCell("1", 2, 2, sheet);
            CreateMockCell("1", 3, 2, sheet);
            CreateMockCell("", 4, 2, sheet);

            CreateMockCell("0", 2, 3, sheet);
            CreateMockCell("-", 3, 3, sheet);
            CreateMockCell("", 4, 3, sheet);

            CreateMockCell("1", 2, 4, sheet);
            CreateMockCell("0", 3, 4, sheet);
            CreateMockCell("0", 4, 4, sheet);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetBuyers();

            Assert.That(result.Count, Is.EqualTo(2));
            Assert.That(result[0].Name, Is.EqualTo("Henrik"));
            Assert.That(result[1].Name, Is.EqualTo("Simpson"));
            Assert.That(result[0].Id, Is.EqualTo(100));
            Assert.That(result[1].Id, Is.EqualTo(101));
        }

        [Test]
        public void SourceReaderGen2_CanGetAmounts()
        {
            /*

                A            B            C            D
            1                Henrik       Alice        Simpson
            2   222222       100          50           75
            3   333333       25           100          200
            4   444444       150          25           300

            */
            var range_A2A4 = CreateMockRange(1, 2, 1, 4);
            var range_B1D1 = CreateMockRange(2, 1, 4, 1);

            var parameters = new InputParameters();
            parameters.ElementIdSpan = range_A2A4.Object;
            parameters.BuyersSpan = range_B1D1.Object;

            // Reusing columns to simplify test
            parameters.BrickLinkDescriptionSpan = range_A2A4.Object;
            parameters.BrickLinkIdSpan = range_A2A4.Object;
            parameters.BrickLinkColorSpan = range_A2A4.Object;
            parameters.TlgColorSpan = range_A2A4.Object;

            var sheet = new Mock<IXLWorksheet>();

            CreateMockCell("Henrik", 1, 2, sheet);
            CreateMockCell("Alice", 1, 3, sheet);
            CreateMockCell("Simpson", 1, 4, sheet);

            CreateMockCell("222222", 2, 1, sheet);
            CreateMockCell("333333", 3, 1, sheet);
            CreateMockCell("444444", 4, 1, sheet);

            CreateMockCell("100", 2, 2, sheet);
            CreateMockCell("25", 3, 2, sheet);
            CreateMockCell("150", 4, 2, sheet);

            CreateMockCell("50", 2, 3, sheet);
            CreateMockCell("100", 3, 3, sheet);
            CreateMockCell("25", 4, 3, sheet);

            CreateMockCell("75", 2, 4, sheet);
            CreateMockCell("200", 3, 4, sheet);
            CreateMockCell("300", 4, 4, sheet);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetReservations();

            Assert.That(result.Count, Is.EqualTo(9));
            Assert.True(result.Any(x => x.Element.ElementID == "222222" && x.Buyer.Name == "Henrik" && x.Amount == 100));
            Assert.True(result.Any(x => x.Element.ElementID == "333333" && x.Buyer.Name == "Henrik" && x.Amount == 25));
            Assert.True(result.Any(x => x.Element.ElementID == "444444" && x.Buyer.Name == "Henrik" && x.Amount == 150));
            Assert.True(result.Any(x => x.Element.ElementID == "222222" && x.Buyer.Name == "Alice" && x.Amount == 50));
            Assert.True(result.Any(x => x.Element.ElementID == "333333" && x.Buyer.Name == "Alice" && x.Amount == 100));
            Assert.True(result.Any(x => x.Element.ElementID == "444444" && x.Buyer.Name == "Alice" && x.Amount == 25));
            Assert.True(result.Any(x => x.Element.ElementID == "222222" && x.Buyer.Name == "Simpson" && x.Amount == 75));
            Assert.True(result.Any(x => x.Element.ElementID == "333333" && x.Buyer.Name == "Simpson" && x.Amount == 200));
            Assert.True(result.Any(x => x.Element.ElementID == "444444" && x.Buyer.Name == "Simpson" && x.Amount == 300));
        }

        [Test]
        public void SourceReaderGen2_WillIgnoreZeroAndNonNumbers()
        {
            /*

                A            B            C            D
            1                Henrik       Alice        Simpson
            2   222222       100          50           
            3   333333       0            100          200
            4   444444       150          -            300

            */
            var range_A2A4 = CreateMockRange(1, 2, 1, 4);
            var range_B1D1 = CreateMockRange(2, 1, 4, 1);

            var parameters = new InputParameters();
            parameters.ElementIdSpan = range_A2A4.Object;
            parameters.BuyersSpan = range_B1D1.Object;

            // Reusing columns to simplify test
            parameters.BrickLinkDescriptionSpan = range_A2A4.Object;
            parameters.BrickLinkIdSpan = range_A2A4.Object;
            parameters.BrickLinkColorSpan = range_A2A4.Object;
            parameters.TlgColorSpan = range_A2A4.Object;

            var sheet = new Mock<IXLWorksheet>();

            CreateMockCell("Henrik", 1, 2, sheet);
            CreateMockCell("Alice", 1, 3, sheet);
            CreateMockCell("Simpson", 1, 4, sheet);

            CreateMockCell("222222", 2, 1, sheet);
            CreateMockCell("333333", 3, 1, sheet);
            CreateMockCell("444444", 4, 1, sheet);

            CreateMockCell("100", 2, 2, sheet);
            CreateMockCell("0", 3, 2, sheet);
            CreateMockCell("150", 4, 2, sheet);

            CreateMockCell("50", 2, 3, sheet);
            CreateMockCell("100", 3, 3, sheet);
            CreateMockCell("-", 4, 3, sheet);

            CreateMockCell("", 2, 4, sheet);
            CreateMockCell("200", 3, 4, sheet);
            CreateMockCell("300", 4, 4, sheet);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetReservations();

            Assert.That(result.Count, Is.EqualTo(6));
            Assert.True(result.Any(x => x.Element.ElementID == "222222" && x.Buyer.Name == "Henrik" && x.Amount == 100));
            Assert.True(result.Any(x => x.Element.ElementID == "444444" && x.Buyer.Name == "Henrik" && x.Amount == 150));
            Assert.True(result.Any(x => x.Element.ElementID == "222222" && x.Buyer.Name == "Alice" && x.Amount == 50));
            Assert.True(result.Any(x => x.Element.ElementID == "333333" && x.Buyer.Name == "Alice" && x.Amount == 100));
            Assert.True(result.Any(x => x.Element.ElementID == "333333" && x.Buyer.Name == "Simpson" && x.Amount == 200));
            Assert.True(result.Any(x => x.Element.ElementID == "444444" && x.Buyer.Name == "Simpson" && x.Amount == 300));
        }
    }
}
