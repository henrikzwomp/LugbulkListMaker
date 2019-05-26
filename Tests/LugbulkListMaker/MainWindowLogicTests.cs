using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LugbulkListMaker;
using ListMakerTwo;
using NUnit.Framework;
using Moq;
using ClosedXML.Excel;
using System.Windows.Media;

namespace Tests.LugbulkListMaker
{
    [TestFixture]
    public class MainWindowLogicTests
    {
        [Test]
        public void SelectingFileWillUpdateSelectFileNameProperty()
        {
            var outside_helper = new Mock<IOutsideWindowHelper>();
            var highlight_worker = new Mock<IHighlightWorker>();

            string selected_file_path = "h:\\henrik.xlsx";

            outside_helper.Setup(x => x.ShowLoadFileDialog(It.IsAny<string>(), out selected_file_path));

            var logic = new MainWindowLogic(outside_helper.Object, highlight_worker.Object);

            Assert.That(logic.SelectFileName, Is.EqualTo("[None]")); // ToDo fix.

            logic.SelectInputFile.Execute(null);

            Assert.That(logic.SelectFileName, Is.EqualTo("henrik.xlsx"));

        }

        [Test]
        public void SelectingFileWillLoadSheetNames()
        {
            var outside_helper = new Mock<IOutsideWindowHelper>();
            var workbook = new Mock<IXLWorkbook>();
            var worksheets = new Mock<IXLWorksheets>();
            var worksheet1 = new Mock<IXLWorksheet>();
            var worksheet2 = new Mock<IXLWorksheet>();
            var cell = new Mock<IXLCell>();
            var address = new Mock<IXLAddress>();
            var highlight_worker = new Mock<IHighlightWorker>();

            string selected_file_path = "h:\\henrik.xlsx";

            outside_helper.Setup(x => x.ShowLoadFileDialog(It.IsAny<string>(), out selected_file_path));
            outside_helper.Setup(x => x.GetXLWorkbook(It.IsAny<string>())).Returns(workbook.Object);
            workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);
            workbook.Setup(x => x.Worksheet(It.IsAny<int>())).Returns(worksheet1.Object);
            worksheets.Setup(x => x.GetEnumerator()).Returns( (new List<IXLWorksheet>() { worksheet1.Object, worksheet2.Object }).GetEnumerator() );
            worksheet1.Setup(x => x.Name).Returns("S1");
            worksheet2.Setup(x => x.Name).Returns("S2");
            worksheet1.Setup(x => x.LastCellUsed()).Returns(cell.Object);
            cell.Setup(x => x.Address).Returns(address.Object);
            address.Setup(x => x.ColumnNumber).Returns(0);
            address.Setup(x => x.RowNumber).Returns(0);
  
            var logic = new MainWindowLogic(outside_helper.Object, highlight_worker.Object);

            Assert.That(logic.SheetNames.Count, Is.EqualTo(0));

            logic.SelectInputFile.Execute(null);

            Assert.That(logic.SheetNames.Count, Is.EqualTo(2));
            Assert.That(logic.SheetNames.Where(x => x == "S1").Count, Is.EqualTo(1));
            Assert.That(logic.SheetNames.Where(x => x == "S2").Count, Is.EqualTo(1));
        }

        [Test]
        public void SelectingFileWillFillDataGridWithData()
        {
            var outside_helper = new Mock<IOutsideWindowHelper>();
            var workbook = new Mock<IXLWorkbook>();
            var worksheets = new Mock<IXLWorksheets>();
            var worksheet1 = new Mock<IXLWorksheet>();
            var cell = new Mock<IXLCell>();
            var address = new Mock<IXLAddress>();
            var highlight_worker = new Mock<IHighlightWorker>();

            string selected_file_path = "h:\\henrik.xlsx";

            outside_helper.Setup(x => x.ShowLoadFileDialog(It.IsAny<string>(), out selected_file_path));
            outside_helper.Setup(x => x.GetXLWorkbook(It.IsAny<string>())).Returns(workbook.Object);
            workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);
            workbook.Setup(x => x.Worksheet(It.IsAny<int>())).Returns(worksheet1.Object);
            worksheets.Setup(x => x.GetEnumerator()).Returns((new List<IXLWorksheet>() { worksheet1.Object }).GetEnumerator());
            worksheet1.Setup(x => x.Name).Returns("S1");

            worksheet1.Setup(x => x.LastCellUsed()).Returns(cell.Object);
            worksheet1.Setup(x => x.Cell(It.IsAny<int>(), It.IsAny<int>())).Returns(cell.Object);
            cell.Setup(x => x.Value).Returns("A Value");
            cell.Setup(x => x.Address).Returns(address.Object);
            address.Setup(x => x.ColumnNumber).Returns(3);
            address.Setup(x => x.RowNumber).Returns(3);

            var logic = new MainWindowLogic(outside_helper.Object, highlight_worker.Object);

            Assert.That(logic.FileData.Count, Is.EqualTo(0));

            logic.SelectInputFile.Execute(null);

            Assert.That(logic.FileData.Count, Is.EqualTo(3));

            /*Assert.That(logic.FileData[0].Count, Is.EqualTo(3 + 1));
            Assert.That(logic.FileData[1].Count, Is.EqualTo(3 + 1));
            Assert.That(logic.FileData[2].Count, Is.EqualTo(3 + 1));
            Assert.That(logic.FileData[1][1], Is.EqualTo("A Value"));*/
            foreach(var item in logic.FileData)
            {
                Assert.That(item.Count(), Is.EqualTo(3 + 1));
            }
            Assert.That(logic.FileData.First().First(), Is.EqualTo("A Value"));
        }

        [Test]
        public void SelectingFileWillUpdateIsFileLoadedValue()
        {
            var outside_helper = new Mock<IOutsideWindowHelper>();
            var workbook = new Mock<IXLWorkbook>();
            var worksheets = new Mock<IXLWorksheets>();
            var worksheet1 = new Mock<IXLWorksheet>();
            var worksheet2 = new Mock<IXLWorksheet>();
            var cell = new Mock<IXLCell>();
            var address = new Mock<IXLAddress>();
            var highlight_worker = new Mock<IHighlightWorker>();

            string selected_file_path = "h:\\henrik.xlsx";

            outside_helper.Setup(x => x.ShowLoadFileDialog(It.IsAny<string>(), out selected_file_path));
            outside_helper.Setup(x => x.GetXLWorkbook(It.IsAny<string>())).Returns(workbook.Object);
            workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);
            workbook.Setup(x => x.Worksheet(It.IsAny<int>())).Returns(worksheet1.Object);
            worksheets.Setup(x => x.GetEnumerator()).Returns((new List<IXLWorksheet>() { worksheet1.Object, worksheet2.Object }).GetEnumerator());
            worksheet1.Setup(x => x.Name).Returns("S1");
            worksheet2.Setup(x => x.Name).Returns("S2");
            worksheet1.Setup(x => x.LastCellUsed()).Returns(cell.Object);
            cell.Setup(x => x.Address).Returns(address.Object);
            address.Setup(x => x.ColumnNumber).Returns(0);
            address.Setup(x => x.RowNumber).Returns(0);

            var logic = new MainWindowLogic(outside_helper.Object, highlight_worker.Object);

            Assert.That(logic.IsFileLoaded, Is.EqualTo(false));

            logic.SelectInputFile.Execute(null);

            Assert.That(logic.IsFileLoaded, Is.EqualTo(true));
        }

        [Test]
        public void ChangingSelectedSheetWillUpdateDataGrid()
        {
            var outside_helper = new Mock<IOutsideWindowHelper>();
            var workbook = new Mock<IXLWorkbook>();
            var worksheets = new Mock<IXLWorksheets>();
            var worksheet1 = new Mock<IXLWorksheet>();
            var worksheet2 = new Mock<IXLWorksheet>();
            var cell1 = new Mock<IXLCell>();
            var cell2 = new Mock<IXLCell>();
            var address1 = new Mock<IXLAddress>();
            var address2 = new Mock<IXLAddress>();
            var highlight_worker = new Mock<IHighlightWorker>();

            string selected_file_path = "h:\\henrik.xlsx";

            outside_helper.Setup(x => x.ShowLoadFileDialog(It.IsAny<string>(), out selected_file_path));
            outside_helper.Setup(x => x.GetXLWorkbook(It.IsAny<string>())).Returns(workbook.Object);
            workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);
            workbook.Setup(x => x.Worksheet(1)).Returns(worksheet1.Object);
            workbook.Setup(x => x.Worksheet(2)).Returns(worksheet2.Object);
            worksheets.Setup(x => x.GetEnumerator()).Returns((new List<IXLWorksheet>() { worksheet1.Object, worksheet2.Object }).GetEnumerator());
            worksheet1.Setup(x => x.Name).Returns("S1");
            worksheet2.Setup(x => x.Name).Returns("S2");
            worksheet1.Setup(x => x.LastCellUsed()).Returns(cell1.Object);
            worksheet1.Setup(x => x.Cell(It.IsAny<int>(), It.IsAny<int>())).Returns(cell2.Object);
            worksheet2.Setup(x => x.LastCellUsed()).Returns(cell2.Object);
            worksheet2.Setup(x => x.Cell(It.IsAny<int>(), It.IsAny<int>())).Returns(cell2.Object);
            cell1.Setup(x => x.Value).Returns("Value 1");
            cell1.Setup(x => x.Address).Returns(address1.Object);
            cell2.Setup(x => x.Value).Returns("Value 2");
            cell2.Setup(x => x.Address).Returns(address2.Object);
            address1.Setup(x => x.ColumnNumber).Returns(0);
            address1.Setup(x => x.RowNumber).Returns(0);
            address2.Setup(x => x.ColumnNumber).Returns(3);
            address2.Setup(x => x.RowNumber).Returns(3);

            var logic = new MainWindowLogic(outside_helper.Object, highlight_worker.Object);

            Assert.That(logic.SheetNames.Count, Is.EqualTo(0));
            Assert.That(logic.SelectedSheetIndex, Is.EqualTo(-1));

            logic.SelectInputFile.Execute(null);

            Assert.That(logic.SheetNames.Count, Is.EqualTo(2));
            Assert.That(logic.SelectedSheetIndex, Is.EqualTo(0));
            Assert.That(logic.FileData.Count, Is.EqualTo(0));

            logic.SelectedSheetIndex = 1;

            Assert.That(logic.FileData.Count, Is.EqualTo(3));
            /*Assert.That(logic.FileData[0].Count, Is.EqualTo(3 + 1));
            Assert.That(logic.FileData[1].Count, Is.EqualTo(3 + 1));
            Assert.That(logic.FileData[2].Count, Is.EqualTo(3 + 1));
            Assert.That(logic.FileData[1][1], Is.EqualTo("Value 2"));*/
            foreach (var item in logic.FileData)
            {
                Assert.That(item.Count(), Is.EqualTo(3 + 1));
            }
            Assert.That(logic.FileData.First().First(), Is.EqualTo("A Value"));

        }

        [Test]
        public void BackgroundWillChangeIfValueIsValidXLAddressOrNot()
        {
            var outside_helper = new Mock<IOutsideWindowHelper>();
            var highlight_worker = new Mock<IHighlightWorker>();

            var workbook = new Mock<IXLWorkbook>();
            var worksheets = new Mock<IXLWorksheets>();
            var worksheet1 = new Mock<IXLWorksheet>();
            var range = ExcelMocker.CreateMockRange(1, 1, 1, 1);

            workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);
            worksheets.Setup(x => x.GetEnumerator()).Returns((new List<IXLWorksheet>() { worksheet1.Object }).GetEnumerator());

            workbook.Setup(x => x.Worksheet(It.IsAny<int>())).Returns(worksheet1.Object);
            worksheet1.Setup(x => x.Range(It.IsAny<string>())).Returns(range.Object);

            outside_helper.Setup(x => x.GetXLWorkbook(It.IsAny<string>())).Returns(workbook.Object);

            string selected_file_path = "h:\\henrik.xlsx";

            outside_helper.Setup(x => x.ShowLoadFileDialog(It.IsAny<string>(), out selected_file_path));
            outside_helper.Setup(x => x.GetXLWorkbook(It.IsAny<string>())).Returns(workbook.Object);

            var cell = ExcelMocker.CreateMockCell("", 1, 1);
            worksheet1.Setup(x => x.LastCellUsed()).Returns(cell.Object);
            worksheet1.Setup(x => x.Cell(It.IsAny<int>(), It.IsAny<int>())).Returns(cell.Object);

            var logic = new MainWindowLogic(outside_helper.Object, highlight_worker.Object);
            logic.SelectInputFile.Execute(null);

            // ElementIdSpan
            Assert.That(logic.ElementIdSpanText, Is.Empty);
            Assert.That(logic.ElementIdSpanBackground.Color.ToString(), Is.EqualTo(Colors.White.ToString()));

            logic.ElementIdSpanText = "A";
            Assert.That(logic.ElementIdSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightPink.ToString()));

            logic.ElementIdSpanText = "A1:C5";
            Assert.That(logic.ElementIdSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightGreen.ToString()));

            // BuyersNamesSpan
            Assert.That(logic.BuyersNamesSpanText, Is.Empty);
            Assert.That(logic.BuyersNamesSpanBackground.Color.ToString(), Is.EqualTo(Colors.White.ToString()));

            logic.BuyersNamesSpanText = "A";
            Assert.That(logic.BuyersNamesSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightPink.ToString()));

            logic.BuyersNamesSpanText = "A1:C5";
            Assert.That(logic.BuyersNamesSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightGreen.ToString()));

            // BlDescSpan
            Assert.That(logic.BlDescSpanText, Is.Empty);
            Assert.That(logic.BlDescSpanBackground.Color.ToString(), Is.EqualTo(Colors.White.ToString()));

            logic.BlDescSpanText = "A";
            Assert.That(logic.BlDescSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightPink.ToString()));

            logic.BlDescSpanText = "A1:C5";
            Assert.That(logic.BlDescSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightGreen.ToString()));

            // BlColorSpan
            Assert.That(logic.BlColorSpanText, Is.Empty);
            Assert.That(logic.BlColorSpanBackground.Color.ToString(), Is.EqualTo(Colors.White.ToString()));

            logic.BlColorSpanText = "A";
            Assert.That(logic.BlColorSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightPink.ToString()));

            logic.BlColorSpanText = "A1:C5";
            Assert.That(logic.BlColorSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightGreen.ToString()));

            // TlgColorSpan
            Assert.That(logic.TlgColorSpanText, Is.Empty);
            Assert.That(logic.TlgColorSpanBackground.Color.ToString(), Is.EqualTo(Colors.White.ToString()));

            logic.TlgColorSpanText = "A";
            Assert.That(logic.TlgColorSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightPink.ToString()));

            logic.TlgColorSpanText = "A1:C5";
            Assert.That(logic.TlgColorSpanBackground.Color.ToString(), Is.EqualTo(Colors.LightGreen.ToString()));


        }
    }
}
