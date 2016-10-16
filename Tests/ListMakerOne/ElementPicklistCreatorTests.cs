﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ListMakerOne;

namespace Tests.ListMakerOne
{
    [TestFixture]
    public class ElementPicklistCreatorTests
    {
        [Test]
        public void CanCreateAListWithCorrectElementValues()
        {
            var reservations = new List<ElementReservation>();
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10001", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10001", Amount = 50 });

            var elements = new List<Element>();
            elements.Add(new Element() { ElementID = "10001", BricklinkDescription = "Plant", BricklinkColor = "Green", MaterialColor = "Dark Green" });

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            Assert.That(picklists.Count, Is.EqualTo(1));
            Assert.That(picklists[0].ElementID, Is.EqualTo("10001"));
            Assert.That(picklists[0].BricklinkDescription, Is.EqualTo("Plant"));
            Assert.That(picklists[0].BricklinkColor, Is.EqualTo("Green"));
            Assert.That(picklists[0].MaterialColor, Is.EqualTo("Dark Green"));
            Assert.That(picklists[0].Reservations.Count, Is.EqualTo(3));
        }

        [Test]
        public void ReservationsWillBeSortedByAmount()
        {
            var reservations = new List<ElementReservation>();
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10001", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10001", Amount = 50 });

            var elements = new List<Element>();
            elements.Add(new Element() { ElementID = "10001", BricklinkDescription = "Plant", BricklinkColor = "Green", MaterialColor = "Dark Green" });

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            Assert.That(picklists.Count, Is.EqualTo(1));
            Assert.That(picklists[0].Reservations.Count, Is.EqualTo(3));
            Assert.That(picklists[0].Reservations[0].Receiver, Is.EqualTo("Alice"));
            Assert.That(picklists[0].Reservations[1].Receiver, Is.EqualTo("Teabox"));
            Assert.That(picklists[0].Reservations[2].Receiver, Is.EqualTo("Simpson"));
        }

        [Test]
        public void ReservationsWillBeSortedByAmountAndReceiver()
        {
            var reservations = new List<ElementReservation>();
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10001", Amount = 100 });

            var elements = new List<Element>();
            elements.Add(new Element() { ElementID = "10001", BricklinkDescription = "Plant", BricklinkColor = "Green", MaterialColor = "Dark Green" });

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            Assert.That(picklists.Count, Is.EqualTo(1));
            Assert.That(picklists[0].Reservations.Count, Is.EqualTo(3));
            Assert.That(picklists[0].Reservations[0].Receiver, Is.EqualTo("Alice"));
            Assert.That(picklists[0].Reservations[1].Receiver, Is.EqualTo("Simpson"));
            Assert.That(picklists[0].Reservations[2].Receiver, Is.EqualTo("Teabox"));
        }

        [Test]
        public void WillOutputListsSortedByElementId()
        {
            var reservations = new List<ElementReservation>();
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10003", Amount = 50 });
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10002", Amount = 200 });

            var elements = new List<Element>();
            elements.Add(new Element() { ElementID = "10002", BricklinkDescription = "Bone", BricklinkColor = "White", MaterialColor = "White" });
            elements.Add(new Element() { ElementID = "10001", BricklinkDescription = "Plant", BricklinkColor = "Green", MaterialColor = "Dark Green" });
            elements.Add(new Element() { ElementID = "10003", BricklinkDescription = "Brick 1 x 2", BricklinkColor = "Dark Red", MaterialColor = "New Dark Red" });

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            Assert.That(picklists.Count, Is.EqualTo(3));
            Assert.That(picklists[0].ElementID, Is.EqualTo("10001"));
            Assert.That(picklists[1].ElementID, Is.EqualTo("10002"));
            Assert.That(picklists[2].ElementID, Is.EqualTo("10003"));
        }

        [Test]
        public void CanHandleSeveralListsCorrectly()
        {
            var reservations = new List<ElementReservation>();
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10003", Amount = 50 });
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10002", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10002", Amount = 1000 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10001", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10003", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10003", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10002", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10001", Amount = 100 });

            var elements = new List<Element>();
            elements.Add(new Element() { ElementID = "10002", BricklinkDescription = "Bone", BricklinkColor = "White", MaterialColor = "White" });
            elements.Add(new Element() { ElementID = "10001", BricklinkDescription = "Plant", BricklinkColor = "Green", MaterialColor = "Dark Green" });
            elements.Add(new Element() { ElementID = "10003", BricklinkDescription = "Brick 1 x 2", BricklinkColor = "Dark Red", MaterialColor = "New Dark Red" });

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            Assert.That(picklists.Count, Is.EqualTo(3));
            Assert.That(picklists[0].ElementID, Is.EqualTo("10001"));
            Assert.That(picklists[1].ElementID, Is.EqualTo("10002"));
            Assert.That(picklists[2].ElementID, Is.EqualTo("10003"));

            Assert.That(picklists[0].Reservations.Count, Is.EqualTo(3));
            Assert.That(picklists[0].Reservations[0].Receiver, Is.EqualTo("Alice"));
            Assert.That(picklists[0].Reservations[1].Receiver, Is.EqualTo("Teabox"));
            Assert.That(picklists[0].Reservations[2].Receiver, Is.EqualTo("Simpson"));

            Assert.That(picklists[1].Reservations.Count, Is.EqualTo(3));
            Assert.That(picklists[1].Reservations[0].Receiver, Is.EqualTo("Alice"));
            Assert.That(picklists[1].Reservations[1].Receiver, Is.EqualTo("Teabox"));
            Assert.That(picklists[1].Reservations[2].Receiver, Is.EqualTo("Simpson"));

            Assert.That(picklists[2].Reservations.Count, Is.EqualTo(3));
            Assert.That(picklists[2].Reservations[0].Receiver, Is.EqualTo("Teabox"));
            Assert.That(picklists[2].Reservations[1].Receiver, Is.EqualTo("Simpson"));
            Assert.That(picklists[2].Reservations[2].Receiver, Is.EqualTo("Alice"));
        }
    }
}
