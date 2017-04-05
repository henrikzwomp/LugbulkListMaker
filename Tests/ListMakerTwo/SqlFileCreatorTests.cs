using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Moq;
using ListMakerTwo;

namespace Tests.ListMakerTwo
{
    [TestFixture]
    public class SqlFileCreatorTests
    {
        [Test]
        public void SqlFileCreator_CanCreateSqls()
        {
            var expected = new List<string>()
            {
                "-- Elements --",
                "INSERT INTO tblElements (ElementId, BlId, Description, BLColor) VALUES (55555, '551b', 'Something', 'Red')",
                "",
                "-- Buyers --",
                "INSERT INTO tblBuyers (Username) VALUES ('Henrik')",
                "",
                "-- Amounts --",
                "INSERT INTO tblBuyersAmounts (Username, ElementId, Amount) VALUES ('Alice', 33333, 200)",
                ""
            };

            var reader = new Mock<ISourceReader>();

            reader.Setup(x => x.GetElements()).Returns(new List<LugBulkElement>() {
                new LugBulkElement() { ElementID = "55555", BricklinkId = "551b",
                    BricklinkDescription = "Something", BricklinkColor = "Red" } });

            reader.Setup(x => x.GetBuyers()).Returns(new List<LugBulkReceiver>() { new LugBulkReceiver() { Name = "Henrik" } });

            reader.Setup(x => x.GetAmounts()).Returns(new List<LugBulkReservation>() {
                new LugBulkReservation() { ElementID = "33333", Receiver = new LugBulkReceiver() { Name = "Alice" }, Amount = 200 }
            });

            var result = SqlFileCreator.MakeFileForLugbulkDatabase(reader.Object);

            Assert.That(result.Count, Is.EqualTo(9));
            Assert.That(result[0], Is.EqualTo(expected[0]));
            Assert.That(result[1], Is.EqualTo(expected[1]));
            Assert.That(result[2], Is.EqualTo(expected[2]));
            Assert.That(result[3], Is.EqualTo(expected[3]));
            Assert.That(result[4], Is.EqualTo(expected[4]));
            Assert.That(result[5], Is.EqualTo(expected[5]));
            Assert.That(result[6], Is.EqualTo(expected[6]));
            Assert.That(result[7], Is.EqualTo(expected[7]));
            Assert.That(result[8], Is.EqualTo(expected[8]));
        }
    }
}
