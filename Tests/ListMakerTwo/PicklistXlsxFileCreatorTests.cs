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
    public class PicklistXlsxFileCreatorTests
    {
        [Test]
        public void PicklistXlsxFileCreator_CanSetElementInfoOnPageOne()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            picklist.BricklinkColor = "Tan";
            picklist.BricklinkDescription = "Brick Brick";
            picklist.ElementID = "1234";
            picklist.MaterialColor = "Brick Yellow";

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            Assert.That(work_sheet.Cell(1, "B").Value.ToString(), Is.EqualTo("1234"));
            Assert.That(work_sheet.Cell(2, "B").Value.ToString(), Is.EqualTo("Brick Brick"));
            Assert.That(work_sheet.Cell(5, "B").Value.ToString(), Is.EqualTo("Tan"));
            Assert.That(work_sheet.Cell(6, "B").Value.ToString(), Is.EqualTo("Brick Yellow"));
        }

        [Test]
        public void PicklistXlsxFileCreator_CanSetElementInfoOnPageTwo()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            picklist.BricklinkColor = "Tan";
            picklist.BricklinkDescription = "Brick Brick";
            picklist.ElementID = "1234";
            picklist.MaterialColor = "Brick Yellow";

            for (int i = 0; i < 60; i++)
                picklist.Reservations.Add(new LugBulkReservation() { Receiver = new LugBulkReceiver() });

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            Assert.That(work_sheet.Cell(57, "B").Value.ToString(), Is.EqualTo("1234"));
            Assert.That(work_sheet.Cell(58, "B").Value.ToString(), Is.EqualTo("Brick Brick"));
            Assert.That(work_sheet.Cell(61, "B").Value.ToString(), Is.EqualTo("Tan"));
            Assert.That(work_sheet.Cell(62, "B").Value.ToString(), Is.EqualTo("Brick Yellow"));
        }

        [Test]
        public void PicklistXlsxFileCreator_CanSetReservationLinesOnPageOne()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            for (int i = 0; i < 30; i++)
                picklist.Reservations.Add(new LugBulkReservation() {
                    Receiver = new LugBulkReceiver() { Id = i, Name = ("B" + i.ToString()) } });

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            for (int i = 0; i < 30; i++)
            {
                Assert.That(work_sheet.Cell(i + 2, "E").Value.ToString(), Is.EqualTo(i.ToString()));
                Assert.That(work_sheet.Cell(i + 2, "F").Value.ToString(), Is.EqualTo("B" + i.ToString()));
            }
        }

        [Test]
        public void PicklistXlsxFileCreator_CanSetReservationLinesOnPageTwo()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            for (int i = 0; i < 90; i++)
                picklist.Reservations.Add(new LugBulkReservation()
                {
                    Receiver = new LugBulkReceiver() { Id = i, Name = ("B" + i.ToString()) }
                });

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            for (int i = 58; i < 90; i++)
            {
                Assert.That(work_sheet.Cell(i + 3, "E").Value.ToString(), Is.EqualTo(i.ToString()));
                Assert.That(work_sheet.Cell(i + 3, "F").Value.ToString(), Is.EqualTo("B" + i.ToString()));
            }
        }

        [Test]
        public void PicklistXlsxFileCreator_WillDeleteReservationLinesOnPageTwoWhenOnlyPageOneIsNeeded()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            for (int i = 0; i < 30; i++)
                picklist.Reservations.Add(new LugBulkReservation()
                {
                    Receiver = new LugBulkReceiver() { Id = i, Name = ("B" + i.ToString()) }
                });

            for (int i = 31; i < 90; i++)
            {
                work_sheet.Cell(i + 2, "E").Value = "X1";
                work_sheet.Cell(i + 2, "F").Value = "X2";
            }

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            for (int i = 31; i < 90; i++)
            {
                Assert.That(work_sheet.Cell(i + 2, "E").Value.ToString(), Is.EqualTo(""));
                Assert.That(work_sheet.Cell(i + 2, "F").Value.ToString(), Is.EqualTo(""));
            }
        }

        [Test]
        public void PicklistXlsxFileCreator_WillDeletePageTwoElementInfoWhenOnlyPageOneIsNeeded()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            picklist.BricklinkColor = "Tan";
            picklist.BricklinkDescription = "Brick Brick";
            picklist.ElementID = "1234";
            picklist.MaterialColor = "Brick Yellow";

            work_sheet.Cell(57, "A").Value = "Junk";
            work_sheet.Cell(58, "A").Value = "Junk";
            work_sheet.Cell(61, "A").Value = "Junk";
            work_sheet.Cell(62, "A").Value = "Junk";
            work_sheet.Cell(57, "B").Value = "Junk";
            work_sheet.Cell(58, "B").Value = "Junk";
            work_sheet.Cell(61, "B").Value = "Junk";
            work_sheet.Cell(62, "B").Value = "Junk";

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            Assert.That(work_sheet.Cell(57, "A").Value.ToString(), Is.EqualTo(""));
            Assert.That(work_sheet.Cell(58, "A").Value.ToString(), Is.EqualTo(""));
            Assert.That(work_sheet.Cell(61, "A").Value.ToString(), Is.EqualTo(""));
            Assert.That(work_sheet.Cell(62, "A").Value.ToString(), Is.EqualTo(""));

            Assert.That(work_sheet.Cell(57, "B").Value.ToString(), Is.EqualTo(""));
            Assert.That(work_sheet.Cell(58, "B").Value.ToString(), Is.EqualTo(""));
            Assert.That(work_sheet.Cell(61, "B").Value.ToString(), Is.EqualTo(""));
            Assert.That(work_sheet.Cell(62, "B").Value.ToString(), Is.EqualTo(""));
        }

        [Test]
        public void PicklistXlsxFileCreator_WillDeletePageTwoMergedCellsWhenOnlyPageOneIsNeeded()
        {
            var work_book = new XLWorkbook();
            work_book.AddWorksheet("TestSheet");
            var work_sheet = work_book.Worksheets.First();
            work_sheet.Range(2, 2, 4, 3).Merge();
            work_sheet.Range(10, 1, 39, 3).Merge();
            work_sheet.Range(58, 2, 60, 3).Merge();
            work_sheet.Range(66, 1, 95, 3).Merge();

            var picklist = new LugBulkPicklist();

            Assert.That(work_sheet.MergedRanges.Count, Is.EqualTo(4));

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 2), Is.True);

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 10), Is.True);

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 58), Is.True);

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 66), Is.True);

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            Assert.That(work_sheet.MergedRanges.Count, Is.EqualTo(2));

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 2), Is.True);

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 10), Is.True);

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 58), Is.False);

            Assert.That(work_sheet.MergedRanges
                .Any(x => x.RangeAddress.FirstAddress.RowNumber == 66), Is.False);
        }
    }
}
