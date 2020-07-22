using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ListMakerTwo;

namespace Tests.ListMakerTwo
{
    [TestFixture]
    public class CsvFileSourceReaderTests
    {
        [Test]
        public void CanGetGetReservations()
        {
            var list_lines = new List<string>()
            {
                "Buyer1,18,4513539,50",
                "Buyer2,19,6247364,100",
                "Buyer3,21,4496343,150",
            };
            var elements_lines = new List<string>() {
                "4513539,Brick1,Dark Blue,Earth Blue,3001",
                "6247364,Brick2,Dark Green,Earth Green,3003",
                "4496343,Brick3,Dark Blue,Earth Blue,3004",
            };

            var reader = new CsvFileSourceReader(list_lines, elements_lines);

            var result = reader.GetReservations();

            Assert.That(result.Count, Is.EqualTo(3));

            Assert.That(result[0].Buyer.Name, Is.EqualTo("Buyer1"));
            Assert.That(result[1].Buyer.Name, Is.EqualTo("Buyer2"));
            Assert.That(result[2].Buyer.Name, Is.EqualTo("Buyer3"));

            Assert.That(result[0].Element.ElementID, Is.EqualTo("4513539"));
            Assert.That(result[1].Element.ElementID, Is.EqualTo("6247364"));
            Assert.That(result[2].Element.ElementID, Is.EqualTo("4496343"));

            Assert.That(result[0].Amount, Is.EqualTo(50));
            Assert.That(result[1].Amount, Is.EqualTo(100));
            Assert.That(result[2].Amount, Is.EqualTo(150));
        }

        [Test]
        public void CanGetBuyers()
        {
            var list_lines = new List<string>()
            {
                "Buyer1,18,4513539,50",
                "Buyer1,19,6247364,100",
                "Buyer2,19,6247364,100",
                "Buyer4,18,4513539,50",
                "Buyer3,19,6247364,100",
                "Buyer3,21,4496343,150",
            };
            var elements_lines = new List<string>() {
                "4513539,Brick1,Dark Blue,Earth Blue,3001",
                "6247364,Brick2,Dark Green,Earth Green,3003",
                "4496343,Brick3,Dark Blue,Earth Blue,3004",
            };

            var reader = new CsvFileSourceReader(list_lines, elements_lines);

            var result = reader.GetBuyers();

            Assert.That(result.Count, Is.EqualTo(4));
            Assert.That(result.Count(x => x.Name == "Buyer1"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.Name == "Buyer2"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.Name == "Buyer3"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.Name == "Buyer4"), Is.EqualTo(1));

            Assert.That(result.First(x => x.Name == "Buyer1").Id, Is.EqualTo(100));
            Assert.That(result.First(x => x.Name == "Buyer2").Id, Is.EqualTo(101));
            Assert.That(result.First(x => x.Name == "Buyer3").Id, Is.EqualTo(102));
            Assert.That(result.First(x => x.Name == "Buyer4").Id, Is.EqualTo(103));
        }

        [Test]
        public void CanGetElements()
        {
            var list_lines = new List<string>()
            {
                "Buyer1,18,4513539,50",
                "Buyer1,19,6247364,100",
                "Buyer2,19,6247364,100",
                "Buyer3,19,6247364,100",
                "Buyer3,21,4496343,150",
                "Buyer4,18,4513539,50",
            };
            var elements_lines = new List<string>() {
                "4513539,Brick1,Dark Blue,Earth Blue,3001",
                "6247364,Brick2,Dark Green,Earth Green,3003",
                "4496343,Brick3,Dark Blue,Earth Blue,3004",
            };

            var reader = new CsvFileSourceReader(list_lines, elements_lines);

            var result = reader.GetElements();

            Assert.That(result.Count, Is.EqualTo(3));
            Assert.That(result.Count(x => x.BricklinkDescription == "Brick1"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.BricklinkDescription == "Brick2"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.BricklinkDescription == "Brick3"), Is.EqualTo(1));

            var brick1 = result.First(x => x.BricklinkDescription == "Brick1");
            var brick2 = result.First(x => x.BricklinkDescription == "Brick2");
            var brick3 = result.First(x => x.BricklinkDescription == "Brick3");

            Assert.That(brick1.BricklinkColor, Is.EqualTo("Dark Blue"));
            Assert.That(brick1.BricklinkId, Is.EqualTo("3001"));
            Assert.That(brick1.ElementID, Is.EqualTo("4513539"));
            Assert.That(brick1.MaterialColor, Is.EqualTo("Earth Blue"));

            Assert.That(brick2.BricklinkColor, Is.EqualTo("Dark Green"));
            Assert.That(brick2.BricklinkId, Is.EqualTo("3003"));
            Assert.That(brick2.ElementID, Is.EqualTo("6247364"));
            Assert.That(brick2.MaterialColor, Is.EqualTo("Earth Green"));

            Assert.That(brick3.BricklinkColor, Is.EqualTo("Dark Blue"));
            Assert.That(brick3.BricklinkId, Is.EqualTo("3004"));
            Assert.That(brick3.ElementID, Is.EqualTo("4496343"));
            Assert.That(brick3.MaterialColor, Is.EqualTo("Earth Blue"));
        }

        [Test]
        public void WillIgnoreTitleRowWhenGettingReservations()
        {
            var list_lines = new List<string>()
            {
                "Nickname,order_id,element_id,antal",
                "Buyer1,18,4513539,50",
                "Buyer2,19,6247364,100",
                "Buyer3,21,4496343,150",
            };
            var elements_lines = new List<string>() {
                "4513539,Brick1,Dark Blue,Earth Blue,3001",
                "6247364,Brick2,Dark Green,Earth Green,3003",
                "4496343,Brick3,Dark Blue,Earth Blue,3004",
            };

            var reader = new CsvFileSourceReader(list_lines, elements_lines);

            var result = reader.GetReservations();

            Assert.That(result.Count, Is.EqualTo(3));

            Assert.That(result[0].Buyer.Name, Is.EqualTo("Buyer1"));
            Assert.That(result[1].Buyer.Name, Is.EqualTo("Buyer2"));
            Assert.That(result[2].Buyer.Name, Is.EqualTo("Buyer3"));

            Assert.That(result[0].Element.ElementID, Is.EqualTo("4513539"));
            Assert.That(result[1].Element.ElementID, Is.EqualTo("6247364"));
            Assert.That(result[2].Element.ElementID, Is.EqualTo("4496343"));

            Assert.That(result[0].Amount, Is.EqualTo(50));
            Assert.That(result[1].Amount, Is.EqualTo(100));
            Assert.That(result[2].Amount, Is.EqualTo(150));
        }

        [Test]
        public void WillIgnoreTitleRowWhenGettingBuyers()
        {
            var list_lines = new List<string>()
            {
                "Nickname,order_id,element_id,antal",
                "Buyer1,18,4513539,50",
                "Buyer1,19,6247364,100",
                "Buyer2,19,6247364,100",
                "Buyer4,18,4513539,50",
                "Buyer3,19,6247364,100",
                "Buyer3,21,4496343,150",
            };
            var elements_lines = new List<string>() {
                "4513539,Brick1,Dark Blue,Earth Blue,3001",
                "6247364,Brick2,Dark Green,Earth Green,3003",
                "4496343,Brick3,Dark Blue,Earth Blue,3004",
            };

            var reader = new CsvFileSourceReader(list_lines, elements_lines);

            var result = reader.GetBuyers();

            Assert.That(result.Count, Is.EqualTo(4));
            Assert.That(result.Count(x => x.Name == "Buyer1"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.Name == "Buyer2"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.Name == "Buyer3"), Is.EqualTo(1));
            Assert.That(result.Count(x => x.Name == "Buyer4"), Is.EqualTo(1));

            Assert.That(result.First(x => x.Name == "Buyer1").Id, Is.EqualTo(100));
            Assert.That(result.First(x => x.Name == "Buyer2").Id, Is.EqualTo(101));
            Assert.That(result.First(x => x.Name == "Buyer3").Id, Is.EqualTo(102));
            Assert.That(result.First(x => x.Name == "Buyer4").Id, Is.EqualTo(103));
        }
    }
}
