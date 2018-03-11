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
    public class SourceReader_RealExample2017
    {
        private static InputParameters GetParameters()
        {
            var _fake_sheet = new XLWorkbook().AddWorksheet("Fake");

            return new InputParameters()
            {
                SourceFileName = TestHelper.AssemblyDirectory + "ListMakerTwo\\LUGBULK 2017 beställning färdig.xlsx",
                WorksheetName = "Beställning sammanställd",
                ElementIdSpan = _fake_sheet.Range("D2:D86"),
                BuyersSpan = _fake_sheet.Range("K87:FW87"),
                BrickLinkDescriptionSpan = _fake_sheet.Range("B2:B86"),
                BrickLinkIdSpan = _fake_sheet.Range("C2:C86"),
                BrickLinkColorSpan = _fake_sheet.Range("E2:E86"),
                TlgColorSpan = _fake_sheet.Range("G2:G86"),
            };
        }

        [Test]
        public void SourceReader_CanGetRightNumberOfBuyersAmount()
        {
            InputParameters parameters = GetParameters();

            var sheet = SheetRetriever.Get(parameters.SourceFileName,
                parameters.WorksheetName);

            var reader = new SourceReader(sheet, parameters);

            var buyers = reader.GetBuyers();

            Assert.That(buyers.Count, Is.GreaterThanOrEqualTo(151));

            Assert.That(buyers.Any(x => x.Name == "Should Not Be Included In Buyers List"), Is.False);

            Assert.That(buyers[0].Name, Is.EqualTo("A"));
            Assert.That(buyers[150].Name, Is.EqualTo("Z"));

            Assert.That(buyers.Count, Is.EqualTo(151));
        }

        
    }
}
