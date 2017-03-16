using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using ListMakerOne;

namespace ListMakerTwo
{
    class Program
    {
        static void Main(string[] args)
        {
            var parameters = new InputParameters()
            {
                SourceFileName = "LUGBULK 2017 beställning färdig.xlsx", 
                WorksheetName = "Beställning sammanställd", 
                ElementRowSpan = "2:86",
                BuyersRow = 87,
                BuyersColumnSpan = "K:FW",
                ElementIdColumn = "D",
                BrickLinkDescriptionColumn = "B",
                BrickLinkIdColumn = "C",
                BrickLinkColorColumn = "E",
                TlgColorColumn = ""
            };

            var sheet = SheetRetriever.Get(parameters.SourceFileName,
                parameters.WorksheetName);

            var reader = new SourceReader(sheet, parameters);


            var picklists = ElementPicklistCreator.CreateLists(reader.GetAmounts(), reader.GetElements());

            Concept03Copy.CreateListFiles(picklists);

            // SqlFileCreator.MakeFileForLugbulkDatabase(reader);
        }
    }
}
