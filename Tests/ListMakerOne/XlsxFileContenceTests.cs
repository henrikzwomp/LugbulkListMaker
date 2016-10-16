using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using NUnit.Framework;
using ListMakerOne;

namespace Tests
{
    [TestFixture]
    public class XlsxFileContenceTests
    {
        [Test]
        public void CanAddCell()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);

            Assert.That(xdoc.ToString().Contains("<row"), Is.False);

            var file = new XlsxFileContence(xdoc);

            file.AddCell(1, 1, "Hello", XlsxFileStyle.None);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);
        }

        [Test]
        public void CanAddCellThatSpanMultibleRowsAndColumns()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);
 

              //<mergeCells count="4">
              // < mergeCell ref= "A1:C2" />

            Assert.That(xdoc.ToString().Contains("<row"), Is.False);
            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"0\""), Is.True);

            var file = new XlsxFileContence(xdoc);

            file.AddCell(1, 1, "Hello", XlsxFileStyle.None, 2, 3);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);

            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"1\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A1:C2\""), Is.True);
        }

        [Test]
        public void CanSetCell()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);

            var file = new XlsxFileContence(xdoc);
            file.AddCell(1, 1, "Hello", XlsxFileStyle.None);
            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);

            file.SetCell(1, 1, "World");

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>World</v>"), Is.True);
        }

        [Test]
        public void CanDeleteCell()
        {
            var worksheet = new XElement("worksheet");
            var sheet_data = new XElement("sheetData");
            worksheet.Add(sheet_data);
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);

            var file = new XlsxFileContence(xdoc);
            file.AddCell(1, 1, "Hello", XlsxFileStyle.None);

            var row = sheet_data.Elements().Where(x => x.Name.LocalName == "row" && x.Attributes().Where(y => y.Name.LocalName == "r" && y.Value == "1").Any()).First();
            var new_col = new XElement("c");
            new_col.Add(new XAttribute("r", "B1"));
            row.Add(new_col);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"B1\""), Is.True);

            file.DeleteCell(1, 1);

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.False);
            Assert.That(xdoc.ToString().Contains("<c r=\"B1\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>World</v>"), Is.False);
        }

        [Test]
        public void DeleteCellWillRemoveRowIfEmpty()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);

            var file = new XlsxFileContence(xdoc);
            file.AddCell(1, 1, "Hello", XlsxFileStyle.None);
            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);

            file.DeleteCell(1, 1);

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.False);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.False);
            Assert.That(xdoc.ToString().Contains("<v>World</v>"), Is.False);
        }

        [Test]
        public void CanDeleteRow()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);

            Assert.That(xdoc.ToString().Contains("<row"), Is.False);

            var file = new XlsxFileContence(xdoc);

            file.AddCell(1, 1, "Hello", XlsxFileStyle.None);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);

            file.DeleteRow(1);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.False);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.False);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.False);
        }

        [Test]
        public void CanRemoveMergeData()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);


            Assert.That(xdoc.ToString().Contains("<row"), Is.False);
            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"0\""), Is.True);

            var file = new XlsxFileContence(xdoc);

            file.AddCell(1, 1, "Hello", XlsxFileStyle.None, 2, 3);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);

            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"1\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A1:C2\""), Is.True);

            file.RemoveMergeData(1, 1);

            Assert.That(xdoc.ToString().Contains("<row r=\"1\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<c r=\"A1\" t=\"str\">"), Is.True);
            Assert.That(xdoc.ToString().Contains("<v>Hello</v>"), Is.True);

            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"0\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"1\""), Is.False);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A1:C2\""), Is.False);
        }

        [Test]
        public void CanRemoveMergeData_UsingDataFromTemplate()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "4"));
            worksheet.Add(mergecells);

            var mergecell1 = new XElement("mergeCell");
            mergecell1.Add(new XAttribute("ref", "A65:C94"));
            mergecells.Add(mergecell1);

            var mergecell2 = new XElement("mergeCell");
            mergecell2.Add(new XAttribute("ref", "A10:C39"));
            mergecells.Add(mergecell2);

            var mergecell3 = new XElement("mergeCell");
            mergecell3.Add(new XAttribute("ref", "B57:C59"));
            mergecells.Add(mergecell3);

            var mergecell4 = new XElement("mergeCell");
            mergecell4.Add(new XAttribute("ref", "B2:C4"));
            mergecells.Add(mergecell4);

            var xdoc = new XDocument(worksheet);

            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"4\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A65:C94\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A10:C39\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"B57:C59\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"B2:C4\""), Is.True);

            var file = new XlsxFileContence(xdoc);

            //Console.WriteLine(xdoc.ToString());

            file.RemoveMergeData(57, 2);

            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"3\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A65:C94\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A10:C39\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"B57:C59\""), Is.False);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"B2:C4\""), Is.True);

            file.RemoveMergeData(65, 1);

            //Console.WriteLine(xdoc.ToString());

            Assert.That(xdoc.ToString().Contains("<mergeCells count=\"2\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A65:C94\""), Is.False);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"A10:C39\""), Is.True);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"B57:C59\""), Is.False);
            Assert.That(xdoc.ToString().Contains("<mergeCell ref=\"B2:C4\""), Is.True);
        }

        [Test]
        public void CanDeleteRowThatDontExists()
        {
            var worksheet = new XElement("worksheet");
            worksheet.Add(new XElement("sheetData"));
            var mergecells = new XElement("mergeCells");
            mergecells.Add(new XAttribute("count", "0"));
            worksheet.Add(mergecells);
            var xdoc = new XDocument(worksheet);

            Assert.That(xdoc.ToString().Contains("<row"), Is.False);

            var file = new XlsxFileContence(xdoc);

            file.AddCell(1, 1, "Hello", XlsxFileStyle.None);

            file.DeleteRow(2);
        }
    }
}
