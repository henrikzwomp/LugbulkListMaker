using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

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

            var picklists = LugBulkPicklistCreator.CreateLists(reader.GetAmounts(), reader.GetElements());

            ///--
            //var picklist = picklists.Where(x => x.ElementID == "6148262").First();
            var picklist = picklists.First();
            File.Copy("Template01.xlsx", picklist.ElementID + ".xlsx", true);

            var workbook = new XLWorkbook(picklist.ElementID + ".xlsx");
            var work_sheet = workbook.Worksheets.First();

            PicklistXlsxFileCreator.Create(work_sheet, picklist);

            workbook.Save();

            ///--

            // SqlFileCreator.MakeFileForLugbulkDatabase(reader);
        }
    }

    
}
