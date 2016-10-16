using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ListMakerOne;
using Moq;

namespace Tests.ListMakerOne
{
    [TestFixture]
    public class TemplateHandlerTests
    {
        [Test]
        public void WillSkipSecondTitleRow()
        {
            var file = new Mock<IXlsxFileContence>();
            var calls = new List<int>();
            file.Setup(x => x.SetCell(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>())).Callback((int z1, int z2, string z3) => { calls.Add(z1); });

            var reservations = new List<ElementReservation>()
            {
                new ElementReservation() { Amount = 10, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 20, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 30, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 40, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 50, ElementID = "1001", Receiver = "Bar" }
            };

            var template_settings = new TemplateSettings();
            template_settings.ReservationsFirstPageEndRow = 4;

            var template_handler = new TemplateHandler(template_settings);


            template_handler.WriteReservations(file.Object, reservations);

            Assert.That(calls, Has.Some.EqualTo(2));
            Assert.That(calls, Has.Some.EqualTo(3));
            Assert.That(calls, Has.Some.EqualTo(4));

            Assert.That(calls, Has.Some.EqualTo(6));
            Assert.That(calls, Has.Some.EqualTo(7));
        }

        [Test]
        public void WillSkipSecondTitleRow2() // Had problem when recievers where one more than what could fit on first page.
        {
            var file = new Mock<IXlsxFileContence>();
            var calls = new List<int>();
            file.Setup(x => x.SetCell(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>())).Callback((int z1, int z2, string z3) => { calls.Add(z1); });

            var reservations = new List<ElementReservation>()
            {
                new ElementReservation() { Amount = 10, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 20, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 30, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 40, ElementID = "1001", Receiver = "Foo" },
            };

            var template_settings = new TemplateSettings();
            template_settings.ReservationsFirstPageEndRow = 4;

            var template_handler = new TemplateHandler(template_settings);

            

            template_handler.WriteReservations(file.Object, reservations);

            Assert.That(calls, Has.Some.EqualTo(2));
            Assert.That(calls, Has.Some.EqualTo(3));
            Assert.That(calls, Has.Some.EqualTo(4));

            Assert.That(calls, Has.Some.EqualTo(6));
        }

        [Test]
        public void WillNotStartDeletingFullRowsUntilAfterRowTen() // After first Information box
        {
            var file = new Mock<IXlsxFileContence>();

            var set_cell_calls = new List<int>();
            var delete_cell_calls = new List<int>();
            var delete_row_calls = new List<int>();

            file.Setup(x => x.SetCell(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>())).Callback((int z1, int z2, string z3) => { set_cell_calls.Add(z1); });
            file.Setup(x => x.DeleteCell(It.IsAny<int>(), It.IsAny<int>())).Callback((int z1, int z2) => { delete_cell_calls.Add(z1); });
            file.Setup(x => x.DeleteRow(It.IsAny<int>())).Callback((int z1) => { delete_row_calls.Add(z1); });

            var reservations = new List<ElementReservation>()
            {
                new ElementReservation() { Amount = 1, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 2, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 3, ElementID = "1001", Receiver = "Alice" },
            };

            var template_settings = new TemplateSettings();
            template_settings.ReservationsFirstPageEndRow = 20;
            template_settings.ReservationsSecondPageEndRow = 150;

            var template_handler = new TemplateHandler(template_settings);

            template_handler.WriteReservations(file.Object, reservations);

            Assert.That(set_cell_calls, Has.Some.EqualTo(2));
            Assert.That(set_cell_calls, Has.Some.EqualTo(3));
            Assert.That(set_cell_calls, Has.Some.EqualTo(4));

            Assert.That(set_cell_calls, Has.None.EqualTo(5));

            Assert.That(delete_cell_calls, Has.Some.EqualTo(5));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(6));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(7));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(8));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(9));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(10));

            Assert.That(delete_row_calls, Has.Some.EqualTo(11));
            Assert.That(delete_row_calls, Has.Some.EqualTo(150));
        }
        
        [Test]
        public void WillNotDeleteTheTenFirstRowOnEachPage() // Will not delete the Information boxes (Fault found in LUGBULK 2015 4211055)
        {
            var file = new Mock<IXlsxFileContence>();

            var delete_cell_calls = new List<int>();
            var delete_row_calls = new List<int>();

            file.Setup(x => x.DeleteCell(It.IsAny<int>(), It.IsAny<int>())).Callback((int z1, int z2) => { delete_cell_calls.Add(z1); });
            file.Setup(x => x.DeleteRow(It.IsAny<int>())).Callback((int z1) => { delete_row_calls.Add(z1); });

            var reservations = new List<ElementReservation>()
            {
                new ElementReservation() { Amount = 1, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 2, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 3, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 4, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 5, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 6, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 7, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 8, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 9, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 10, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 11, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 12, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 13, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 14, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 15, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 16, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 17, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 18, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 19, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 20, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 21, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 22, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 23, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 24, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 25, ElementID = "1001", Receiver = "Bar" },
            };

            var template_settings = new TemplateSettings();
            template_settings.ReservationsFirstPageEndRow = 20;
            template_settings.ReservationsSecondPageEndRow = 150;

            var template_handler = new TemplateHandler(template_settings);

            template_handler.WriteReservations(file.Object, reservations);

            Assert.That(delete_row_calls, Has.None.EqualTo(1));
            Assert.That(delete_row_calls, Has.None.EqualTo(2)); // Reservation #1
            Assert.That(delete_row_calls, Has.None.EqualTo(3));
            Assert.That(delete_row_calls, Has.None.EqualTo(4));
            Assert.That(delete_row_calls, Has.None.EqualTo(5));
            Assert.That(delete_row_calls, Has.None.EqualTo(6));
            Assert.That(delete_row_calls, Has.None.EqualTo(7));
            Assert.That(delete_row_calls, Has.None.EqualTo(8));
            Assert.That(delete_row_calls, Has.None.EqualTo(9));
            Assert.That(delete_row_calls, Has.None.EqualTo(10));
            Assert.That(delete_row_calls, Has.None.EqualTo(11));
            Assert.That(delete_row_calls, Has.None.EqualTo(12));
            Assert.That(delete_row_calls, Has.None.EqualTo(13));
            Assert.That(delete_row_calls, Has.None.EqualTo(14));
            Assert.That(delete_row_calls, Has.None.EqualTo(15));
            Assert.That(delete_row_calls, Has.None.EqualTo(16));
            Assert.That(delete_row_calls, Has.None.EqualTo(17));
            Assert.That(delete_row_calls, Has.None.EqualTo(18));
            Assert.That(delete_row_calls, Has.None.EqualTo(19));
            Assert.That(delete_row_calls, Has.None.EqualTo(20)); // Reservation #19

            Assert.That(delete_row_calls, Has.None.EqualTo(22)); // Reservation #20
            Assert.That(delete_row_calls, Has.None.EqualTo(23));
            Assert.That(delete_row_calls, Has.None.EqualTo(24));
            Assert.That(delete_row_calls, Has.None.EqualTo(25));
            Assert.That(delete_row_calls, Has.None.EqualTo(26));
            Assert.That(delete_row_calls, Has.None.EqualTo(27)); // Reservation #25

            Assert.That(delete_cell_calls, Has.Some.EqualTo(28));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(29));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(30));

            Assert.That(delete_row_calls, Has.Some.EqualTo(31));
            Assert.That(delete_row_calls, Has.Some.EqualTo(32));
            Assert.That(delete_row_calls, Has.Some.EqualTo(33));
            Assert.That(delete_row_calls, Has.Some.EqualTo(34));
            Assert.That(delete_row_calls, Has.Some.EqualTo(35));
            Assert.That(delete_row_calls, Has.Some.EqualTo(50));
            Assert.That(delete_row_calls, Has.Some.EqualTo(150));
        }

        [Test]
        public void WillRemoveSecondPageTitleRowIfRecieversEqualToFirstPageEndRowMinusOne() // Will remove second page title row if recievers equal = ReservationsSecondPageEndRow - 1 (Fault found in LUGBULK 2016 (First faulty version) 4107432)
        {
            var file = new Mock<IXlsxFileContence>();

            var delete_row_calls = new List<int>();

            file.Setup(x => x.DeleteRow(It.IsAny<int>())).Callback((int z1) => { delete_row_calls.Add(z1); });

            var reservations = new List<ElementReservation>()
            {
                new ElementReservation() { Amount = 1, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 2, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 3, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 4, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 5, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 6, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 7, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 8, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 9, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 10, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 11, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 12, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 13, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 14, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 15, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 16, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 17, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 18, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 19, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 20, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 21, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 22, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 23, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 24, ElementID = "1001", Receiver = "Foo" }
            };

            var template_settings = new TemplateSettings();
            template_settings.ReservationsFirstPageEndRow = 25;
            template_settings.ReservationsSecondPageEndRow = 100;

            var template_handler = new TemplateHandler(template_settings);

            template_handler.WriteReservations(file.Object, reservations);

            Assert.That(delete_row_calls, Has.None.EqualTo(25));// Reservation #25
            Assert.That(delete_row_calls, Has.Some.EqualTo(27));
            Assert.That(delete_row_calls, Has.Some.EqualTo(26));
        }

        [Test]
        public void WillNotDeleteTheTenFirstRowOnEachPageIfRecieversEqualToFirstPageEndRowPlusOne() // Fault found in LUGBULK 2016 (First faulty version) 4211055
        {
            var file = new Mock<IXlsxFileContence>();

            var set_cell_calls = new List<int>();
            var delete_cell_calls = new List<int>();
            var delete_row_calls = new List<int>();

            file.Setup(x => x.SetCell(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>())).Callback((int z1, int z2, string z3) => { set_cell_calls.Add(z1); });
            file.Setup(x => x.DeleteCell(It.IsAny<int>(), It.IsAny<int>())).Callback((int z1, int z2) => { delete_cell_calls.Add(z1); });
            file.Setup(x => x.DeleteRow(It.IsAny<int>())).Callback((int z1) => { delete_row_calls.Add(z1); });

            var reservations = new List<ElementReservation>()
            {
                new ElementReservation() { Amount = 1, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 2, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 3, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 4, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 5, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 6, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 7, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 8, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 9, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 10, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 11, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 12, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 13, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 14, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 15, ElementID = "1001", Receiver = "Bar" },
                new ElementReservation() { Amount = 16, ElementID = "1001", Receiver = "Henrik" },
                new ElementReservation() { Amount = 17, ElementID = "1001", Receiver = "Simpson" },
                new ElementReservation() { Amount = 18, ElementID = "1001", Receiver = "Alice" },
                new ElementReservation() { Amount = 19, ElementID = "1001", Receiver = "Foo" },
                new ElementReservation() { Amount = 20, ElementID = "1001", Receiver = "Bar" }
            };

            var template_settings = new TemplateSettings();
            template_settings.ReservationsFirstPageEndRow = 20;
            template_settings.ReservationsSecondPageEndRow = 150;

            var template_handler = new TemplateHandler(template_settings);

            template_handler.WriteReservations(file.Object, reservations);

            Assert.That(delete_row_calls, Has.None.EqualTo(1));
            Assert.That(delete_row_calls, Has.None.EqualTo(2)); // Reservation #1
            Assert.That(delete_row_calls, Has.None.EqualTo(3));
            Assert.That(delete_row_calls, Has.None.EqualTo(4));
            Assert.That(delete_row_calls, Has.None.EqualTo(5));
            Assert.That(delete_row_calls, Has.None.EqualTo(6));
            Assert.That(delete_row_calls, Has.None.EqualTo(7));
            Assert.That(delete_row_calls, Has.None.EqualTo(8));
            Assert.That(delete_row_calls, Has.None.EqualTo(9));
            Assert.That(delete_row_calls, Has.None.EqualTo(10));
            Assert.That(delete_row_calls, Has.None.EqualTo(11));
            Assert.That(delete_row_calls, Has.None.EqualTo(12));
            Assert.That(delete_row_calls, Has.None.EqualTo(13));
            Assert.That(delete_row_calls, Has.None.EqualTo(14));
            Assert.That(delete_row_calls, Has.None.EqualTo(15));
            Assert.That(delete_row_calls, Has.None.EqualTo(16));
            Assert.That(delete_row_calls, Has.None.EqualTo(17));
            Assert.That(delete_row_calls, Has.None.EqualTo(18));
            Assert.That(delete_row_calls, Has.None.EqualTo(19));
            Assert.That(delete_row_calls, Has.None.EqualTo(20)); // Reservation #19

            Assert.That(delete_row_calls, Has.None.EqualTo(22)); // Reservation #20
            Assert.That(delete_row_calls, Has.None.EqualTo(23));
            Assert.That(delete_row_calls, Has.None.EqualTo(24));
            Assert.That(delete_row_calls, Has.None.EqualTo(25));

            Assert.That(delete_cell_calls, Has.Some.EqualTo(23));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(24));
            Assert.That(delete_cell_calls, Has.Some.EqualTo(25));

            Assert.That(delete_row_calls, Has.Some.EqualTo(31));
            Assert.That(delete_row_calls, Has.Some.EqualTo(32));
            Assert.That(delete_row_calls, Has.Some.EqualTo(33));
            Assert.That(delete_row_calls, Has.Some.EqualTo(34));
            Assert.That(delete_row_calls, Has.Some.EqualTo(35));
            Assert.That(delete_row_calls, Has.Some.EqualTo(50));
            Assert.That(delete_row_calls, Has.Some.EqualTo(150));
        }
    }
}
