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
        [Test]
        public void SourceReader_CanGetElements()
        {
            var parameters = new InputParameters();

            parameters.ElementRowSpan = "3:5";
            parameters.ElementIdColumn = "E";
            parameters.BrickLinkDescriptionColumn = "D";
            parameters.BrickLinkIdColumn = "I";
            parameters.BrickLinkColorColumn = "C";

            var sheet = new Mock<IXLWorksheet>();

            var cell_E3 = new Mock<IXLCell>(); cell_E3.SetupGet(x => x.Value).Returns("3333");
            sheet.Setup(x => x.Cell(3, "E")).Returns(cell_E3.Object);
            var cell_D3 = new Mock<IXLCell>(); cell_D3.SetupGet(x => x.Value).Returns("Brick 3");
            sheet.Setup(x => x.Cell(3, "D")).Returns(cell_D3.Object);
            var cell_I3 = new Mock<IXLCell>(); cell_I3.SetupGet(x => x.Value).Returns("3333b");
            sheet.Setup(x => x.Cell(3, "I")).Returns(cell_I3.Object);
            var cell_C3 = new Mock<IXLCell>(); cell_C3.SetupGet(x => x.Value).Returns("Red");
            sheet.Setup(x => x.Cell(3, "C")).Returns(cell_C3.Object);

            var cell_E4 = new Mock<IXLCell>(); cell_E4.SetupGet(x => x.Value).Returns("4444");
            sheet.Setup(x => x.Cell(4, "E")).Returns(cell_E4.Object);
            var cell_D4 = new Mock<IXLCell>(); cell_D4.SetupGet(x => x.Value).Returns("Brick 4");
            sheet.Setup(x => x.Cell(4, "D")).Returns(cell_D4.Object);
            var cell_I4 = new Mock<IXLCell>(); cell_I4.SetupGet(x => x.Value).Returns("4444b");
            sheet.Setup(x => x.Cell(4, "I")).Returns(cell_I4.Object);
            var cell_C4 = new Mock<IXLCell>(); cell_C4.SetupGet(x => x.Value).Returns("Blue");
            sheet.Setup(x => x.Cell(4, "C")).Returns(cell_C4.Object);

            var cell_E5 = new Mock<IXLCell>(); cell_E5.SetupGet(x => x.Value).Returns("5555");
            sheet.Setup(x => x.Cell(5, "E")).Returns(cell_E5.Object);
            var cell_D5 = new Mock<IXLCell>(); cell_D5.SetupGet(x => x.Value).Returns("Brick 5");
            sheet.Setup(x => x.Cell(5, "D")).Returns(cell_D5.Object);
            var cell_I5 = new Mock<IXLCell>(); cell_I5.SetupGet(x => x.Value).Returns("5555b");
            sheet.Setup(x => x.Cell(5, "I")).Returns(cell_I5.Object);
            var cell_C5 = new Mock<IXLCell>(); cell_C5.SetupGet(x => x.Value).Returns("Green");
            sheet.Setup(x => x.Cell(5, "C")).Returns(cell_C5.Object);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetElements();

            Assert.That(result.Count, Is.EqualTo(3));
            Assert.That(result[0].ElementID, Is.EqualTo("3333"));
            Assert.That(result[0].BricklinkDescription, Is.EqualTo("Brick 3"));
            Assert.That(result[0].BricklinkId, Is.EqualTo("3333b"));
            Assert.That(result[0].BricklinkColor, Is.EqualTo("Red"));
            Assert.That(result[1].ElementID, Is.EqualTo("4444"));
            Assert.That(result[1].BricklinkDescription, Is.EqualTo("Brick 4"));
            Assert.That(result[1].BricklinkId, Is.EqualTo("4444b"));
            Assert.That(result[1].BricklinkColor, Is.EqualTo("Blue"));
            Assert.That(result[2].ElementID, Is.EqualTo("5555"));
            Assert.That(result[2].BricklinkDescription, Is.EqualTo("Brick 5"));
            Assert.That(result[2].BricklinkId, Is.EqualTo("5555b"));
            Assert.That(result[2].BricklinkColor, Is.EqualTo("Green"));

            // if (amount_cell.Address.ColumnLetter == Buyers_Column_Span_End)
        }

        [Test]
        public void SourceReader_CanGetBuyers()
        {
            var parameters = new InputParameters();

            parameters.BuyersColumnSpan = "C:E";
            parameters.BuyersRow = 2;

            var sheet = new Mock<IXLWorksheet>();

            var cell_C2 = new Mock<IXLCell>(); cell_C2.SetupGet(x => x.Value).Returns("Henrik");
            var cell_C2_address = new Mock<IXLAddress>(); cell_C2_address.SetupGet(x => x.ColumnLetter).Returns("C");
            cell_C2.SetupGet(x => x.Address).Returns(cell_C2_address.Object);

            var cell_D2 = new Mock<IXLCell>(); cell_D2.SetupGet(x => x.Value).Returns("Alice");
            var cell_D2_address = new Mock<IXLAddress>(); cell_D2_address.SetupGet(x => x.ColumnLetter).Returns("D");
            cell_D2.SetupGet(x => x.Address).Returns(cell_D2_address.Object);

            var cell_E2 = new Mock<IXLCell>(); cell_E2.SetupGet(x => x.Value).Returns("Simpson");
            var cell_E2_address = new Mock<IXLAddress>(); cell_E2_address.SetupGet(x => x.ColumnLetter).Returns("E");
            cell_E2.SetupGet(x => x.Address).Returns(cell_E2_address.Object);
            
            sheet.Setup(x => x.Cell(2, "C")).Returns(cell_C2.Object);

            cell_C2.Setup(x => x.CellRight()).Returns(cell_D2.Object);
            cell_D2.Setup(x => x.CellRight()).Returns(cell_E2.Object);

            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetBuyers();

            Assert.That(result.Count, Is.EqualTo(3));
            Assert.That(result[0].Name, Is.EqualTo("Henrik"));
            Assert.That(result[1].Name, Is.EqualTo("Alice"));
            Assert.That(result[2].Name, Is.EqualTo("Simpson"));
        }

        [Test]
        public void SourceReader_CanGetAmounts()
        {
            /*

                A            B            C            D
            1                Henrik       Alice        Simpson
            2   222222       100          50             
            3   333333                    100          200
            4   444444       150                       300

            */

            var parameters = new InputParameters();
            parameters.ElementRowSpan = "2:4";
            parameters.ElementIdColumn = "A";
            parameters.BuyersRow = 1;
            parameters.BuyersColumnSpan = "B:D";

            var cell_A1 = new Mock<IXLCell>(); cell_A1.SetupGet(x => x.Value).Returns("");
            var cell_A2 = new Mock<IXLCell>(); cell_A2.SetupGet(x => x.Value).Returns("222222");
            var cell_A3 = new Mock<IXLCell>(); cell_A3.SetupGet(x => x.Value).Returns("333333");
            var cell_A4 = new Mock<IXLCell>(); cell_A4.SetupGet(x => x.Value).Returns("444444");

            var cell_B1 = new Mock<IXLCell>(); cell_B1.SetupGet(x => x.Value).Returns("Henrik");
            var cell_B2 = new Mock<IXLCell>(); cell_B2.SetupGet(x => x.Value).Returns("100");
            var cell_B3 = new Mock<IXLCell>(); cell_B3.SetupGet(x => x.Value).Returns("");
            var cell_B4 = new Mock<IXLCell>(); cell_B4.SetupGet(x => x.Value).Returns("150");

            var cell_C1 = new Mock<IXLCell>(); cell_C1.SetupGet(x => x.Value).Returns("Alice");
            var cell_C2 = new Mock<IXLCell>(); cell_C2.SetupGet(x => x.Value).Returns("50");
            var cell_C3 = new Mock<IXLCell>(); cell_C3.SetupGet(x => x.Value).Returns("100");
            var cell_C4 = new Mock<IXLCell>(); cell_C4.SetupGet(x => x.Value).Returns("");

            var cell_D1 = new Mock<IXLCell>(); cell_D1.SetupGet(x => x.Value).Returns("Simpson");
            var cell_D2 = new Mock<IXLCell>(); cell_D2.SetupGet(x => x.Value).Returns("");
            var cell_D3 = new Mock<IXLCell>(); cell_D3.SetupGet(x => x.Value).Returns("200");
            var cell_D4 = new Mock<IXLCell>(); cell_D4.SetupGet(x => x.Value).Returns("300");

            cell_B1.Setup(x => x.CellRight()).Returns(cell_C1.Object);
            cell_C1.Setup(x => x.CellRight()).Returns(cell_D1.Object);
            cell_B2.Setup(x => x.CellRight()).Returns(cell_C2.Object);
            cell_C2.Setup(x => x.CellRight()).Returns(cell_D2.Object);
            cell_B3.Setup(x => x.CellRight()).Returns(cell_C3.Object);
            cell_C3.Setup(x => x.CellRight()).Returns(cell_D3.Object);
            cell_B4.Setup(x => x.CellRight()).Returns(cell_C4.Object);
            cell_C4.Setup(x => x.CellRight()).Returns(cell_D4.Object);

            var B_address = new Mock<IXLAddress>(); B_address.SetupGet(x => x.ColumnLetter).Returns("B");
            cell_B1.SetupGet(x => x.Address).Returns(B_address.Object);
            cell_B2.SetupGet(x => x.Address).Returns(B_address.Object);
            cell_B3.SetupGet(x => x.Address).Returns(B_address.Object);
            cell_B4.SetupGet(x => x.Address).Returns(B_address.Object);

            var C_address = new Mock<IXLAddress>(); C_address.SetupGet(x => x.ColumnLetter).Returns("C");
            cell_C1.SetupGet(x => x.Address).Returns(C_address.Object);
            cell_C2.SetupGet(x => x.Address).Returns(C_address.Object);
            cell_C3.SetupGet(x => x.Address).Returns(C_address.Object);
            cell_C4.SetupGet(x => x.Address).Returns(C_address.Object);

            var D_address = new Mock<IXLAddress>(); D_address.SetupGet(x => x.ColumnLetter).Returns("D");
            cell_D1.SetupGet(x => x.Address).Returns(D_address.Object);
            cell_D2.SetupGet(x => x.Address).Returns(D_address.Object);
            cell_D3.SetupGet(x => x.Address).Returns(D_address.Object);
            cell_D4.SetupGet(x => x.Address).Returns(D_address.Object);

            var sheet = new Mock<IXLWorksheet>();

            sheet.Setup(x => x.Cell(2, parameters.ElementIdColumn)).Returns(cell_A2.Object);
            sheet.Setup(x => x.Cell(3, parameters.ElementIdColumn)).Returns(cell_A3.Object);
            sheet.Setup(x => x.Cell(4, parameters.ElementIdColumn)).Returns(cell_A4.Object);

            sheet.Setup(x => x.Cell(2, "B")).Returns(cell_B2.Object);
            sheet.Setup(x => x.Cell(3, "B")).Returns(cell_B3.Object);
            sheet.Setup(x => x.Cell(4, "B")).Returns(cell_B4.Object);

            sheet.Setup(x => x.Cell(parameters.BuyersRow, "B")).Returns(cell_B1.Object);


            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetAmounts();
            Assert.That(result.Count, Is.EqualTo(6));
            Assert.True(result.Any(x => x.ElementID == "222222" && x.Receiver.Name == "Henrik" && x.Amount == 100));
            Assert.True(result.Any(x => x.ElementID == "444444" && x.Receiver.Name == "Henrik" && x.Amount == 150));
            Assert.True(result.Any(x => x.ElementID == "222222" && x.Receiver.Name == "Alice" && x.Amount == 50));
            Assert.True(result.Any(x => x.ElementID == "333333" && x.Receiver.Name == "Alice" && x.Amount == 100));
            Assert.True(result.Any(x => x.ElementID == "333333" && x.Receiver.Name == "Simpson" && x.Amount == 200));
            Assert.True(result.Any(x => x.ElementID == "444444" && x.Receiver.Name == "Simpson" && x.Amount == 300));
        }

        [Test]
        public void SourceReader_WillIgnoreZeroAndNonNumbers()
        {
            /*

                A            B            C            
            1                Henrik       Alice        
            2   222222       100          -
            3   333333       0            100          

            */

            var parameters = new InputParameters();
            parameters.ElementRowSpan = "2:3";
            parameters.ElementIdColumn = "A";
            parameters.BuyersRow = 1;
            parameters.BuyersColumnSpan = "B:C";

            var cell_A1 = new Mock<IXLCell>(); cell_A1.SetupGet(x => x.Value).Returns("");
            var cell_A2 = new Mock<IXLCell>(); cell_A2.SetupGet(x => x.Value).Returns("222222");
            var cell_A3 = new Mock<IXLCell>(); cell_A3.SetupGet(x => x.Value).Returns("333333");

            var cell_B1 = new Mock<IXLCell>(); cell_B1.SetupGet(x => x.Value).Returns("Henrik");
            var cell_B2 = new Mock<IXLCell>(); cell_B2.SetupGet(x => x.Value).Returns("100");
            var cell_B3 = new Mock<IXLCell>(); cell_B3.SetupGet(x => x.Value).Returns("0");

            var cell_C1 = new Mock<IXLCell>(); cell_C1.SetupGet(x => x.Value).Returns("Alice");
            var cell_C2 = new Mock<IXLCell>(); cell_C2.SetupGet(x => x.Value).Returns("-");
            var cell_C3 = new Mock<IXLCell>(); cell_C3.SetupGet(x => x.Value).Returns("100");

            cell_B1.Setup(x => x.CellRight()).Returns(cell_C1.Object);
            cell_B2.Setup(x => x.CellRight()).Returns(cell_C2.Object);
            cell_B3.Setup(x => x.CellRight()).Returns(cell_C3.Object);

            var B_address = new Mock<IXLAddress>(); B_address.SetupGet(x => x.ColumnLetter).Returns("B");
            cell_B1.SetupGet(x => x.Address).Returns(B_address.Object);
            cell_B2.SetupGet(x => x.Address).Returns(B_address.Object);
            cell_B3.SetupGet(x => x.Address).Returns(B_address.Object);

            var C_address = new Mock<IXLAddress>(); C_address.SetupGet(x => x.ColumnLetter).Returns("C");
            cell_C1.SetupGet(x => x.Address).Returns(C_address.Object);
            cell_C2.SetupGet(x => x.Address).Returns(C_address.Object);
            cell_C3.SetupGet(x => x.Address).Returns(C_address.Object);

            var sheet = new Mock<IXLWorksheet>();

            sheet.Setup(x => x.Cell(2, parameters.ElementIdColumn)).Returns(cell_A2.Object);
            sheet.Setup(x => x.Cell(3, parameters.ElementIdColumn)).Returns(cell_A3.Object);

            sheet.Setup(x => x.Cell(2, "B")).Returns(cell_B2.Object);
            sheet.Setup(x => x.Cell(3, "B")).Returns(cell_B3.Object);

            sheet.Setup(x => x.Cell(parameters.BuyersRow, "B")).Returns(cell_B1.Object);


            var reader = new SourceReader(sheet.Object, parameters);

            var result = reader.GetAmounts();
            Assert.That(result.Count, Is.EqualTo(2));
            Assert.True(result.Any(x => x.ElementID == "222222" && x.Receiver.Name == "Henrik" && x.Amount == 100));
            Assert.True(result.Any(x => x.ElementID == "333333" && x.Receiver.Name == "Alice" && x.Amount == 100));
        }
    }
}
