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
            var _fake_sheet = new XLWorkbook().AddWorksheet("Fake");

            var parameters = new InputParameters()
            {
                SourceFileName = "LUGBULK 2017 beställning färdig.xlsx",
                WorksheetName = "Beställning sammanställd",
                ElementIdSpan = _fake_sheet.Range("D2:D86"),
                BuyersSpan = _fake_sheet.Range("K87:FW87"),
                BrickLinkDescriptionSpan = _fake_sheet.Range("B2:B86"),
                BrickLinkIdSpan = _fake_sheet.Range("C2:C86"),
                BrickLinkColorSpan = _fake_sheet.Range("E2:E86"),
                TlgColorSpan = _fake_sheet.Range("G2:G86"),
            };

            var sheet = SheetRetriever.Get(parameters.SourceFileName,
                parameters.WorksheetName);

            var reader = new SourceReader(sheet, parameters);

            var base_out_put_folder = "D:\\Henrik\\LEGO\\LUGBULK2017";

            if (!Directory.Exists(base_out_put_folder))
                Directory.CreateDirectory(base_out_put_folder);

            Console.WriteLine("Creating list of Buyers...");
            CreateBuyersList(reader, base_out_put_folder);
            Console.WriteLine("Creating Picklists...");
            CreatePicklists(reader, base_out_put_folder);
            Console.WriteLine("Creating lists of buyers reservations...");
            CreateBuyerSummeryLists(reader, base_out_put_folder);
            Console.WriteLine("Creating a Master list...");
            CreateMasterlist(reader, base_out_put_folder);

            // SqlFileCreator.MakeFileForLugbulkDatabase(reader);
        }

        private static void CreateMasterlist(SourceReader reader, string base_out_put_folder)
        {
            var output_file = base_out_put_folder + "\\MasterList.xlsx";

            var reservations = reader.GetReservations();

            var workbook = new XLWorkbook();
            var work_sheet = workbook.AddWorksheet("Master List");

            work_sheet.Cell(1, "A").Value = "Buyer.Id"; work_sheet.Cell(1, "A").Style.Font.Bold = true;
            work_sheet.Cell(1, "B").Value = "Buyer.Name"; work_sheet.Cell(1, "B").Style.Font.Bold = true;
            work_sheet.Cell(1, "C").Value = "Element.ElementID"; work_sheet.Cell(1, "C").Style.Font.Bold = true;
            work_sheet.Cell(1, "D").Value = "Element.BricklinkId"; work_sheet.Cell(1, "D").Style.Font.Bold = true;
            work_sheet.Cell(1, "E").Value = "Element.BricklinkDescription"; work_sheet.Cell(1, "E").Style.Font.Bold = true;
            work_sheet.Cell(1, "F").Value = "Element.BricklinkColor"; work_sheet.Cell(1, "F").Style.Font.Bold = true;
            work_sheet.Cell(1, "G").Value = "Element.MaterialColor"; work_sheet.Cell(1, "G").Style.Font.Bold = true;
            work_sheet.Cell(1, "H").Value = "Amount"; work_sheet.Cell(1, "H").Style.Font.Bold = true;

            int line_count = 2;
            foreach (var reservation in reservations)
            {
                work_sheet.Cell(line_count, "A").Value = reservation.Buyer.Id;
                work_sheet.Cell(line_count, "B").Value = reservation.Buyer.Name;
                work_sheet.Cell(line_count, "C").Value = reservation.Element.ElementID;
                work_sheet.Cell(line_count, "D").Value = reservation.Element.BricklinkId;
                work_sheet.Cell(line_count, "E").Value = reservation.Element.BricklinkDescription;
                work_sheet.Cell(line_count, "F").Value = reservation.Element.BricklinkColor;
                work_sheet.Cell(line_count, "G").Value = reservation.Element.MaterialColor;
                work_sheet.Cell(line_count, "H").Value = reservation.Amount;
                line_count++;
            }

            if (File.Exists(output_file))
                File.Delete(output_file);

            workbook.SaveAs(output_file);
        }

        private static void CreateBuyerSummeryLists(SourceReader reader, string base_out_put_folder)
        {
            var buyers = reader.GetBuyers();
            var reservations = reader.GetReservations();

            var output_folder = base_out_put_folder + "\\BuyerSummeryLists";

            if (!Directory.Exists(output_folder))
                Directory.CreateDirectory(output_folder);

            foreach (var buyer in buyers)
            {
                var new_file_path = output_folder + "\\" + buyer.Id + "-" + buyer.Name + ".xlsx";

                File.Copy("Templates\\ReceiverSummeryTemplate01.xlsx", new_file_path, true);

                var workbook = new XLWorkbook(new_file_path);
                var work_sheet = workbook.Worksheets.First();

                BuyerSummeryFileCreator.Create(work_sheet, buyer, 
                    reservations.Where(x => x.Buyer.Id == buyer.Id));

                workbook.Save();
            }
        }

        private static void CreateBuyersList(SourceReader reader, string base_out_put_folder)
        {
            File.Copy("Templates\\ReceiversListTemplate01.xlsx", base_out_put_folder + "\\BuyersList.xlsx", true);

            var workbook = new XLWorkbook(base_out_put_folder + "\\BuyersList.xlsx");
            var work_sheet = workbook.Worksheets.First();

            var buyers = reader.GetBuyers();

            for(int i = 0; i < buyers.Count; i++)
            {
                work_sheet.Cell(i + 2, "A").Value = buyers[i].Id;
                work_sheet.Cell(i + 2, "A").Style.Alignment
                    .SetHorizontal(XLAlignmentHorizontalValues.Center);
                work_sheet.Cell(i + 2, "B").Value = buyers[i].Name;
            }

            workbook.Save();
        }

        private static void CreatePicklists(SourceReader reader, string base_out_put_folder)
        {
            var picklists = LugBulkPicklistCreator.CreateLists(reader.GetReservations(), reader.GetElements());

            var output_folder = base_out_put_folder + "\\Picklists\\";

            if (!Directory.Exists(output_folder))
                Directory.CreateDirectory(output_folder);

            foreach (var picklist in picklists)
            {
                File.Copy("Templates\\PicklistTemplate01.xlsx", output_folder + picklist.ElementID + ".xlsx", true);

                var workbook = new XLWorkbook(output_folder + picklist.ElementID + ".xlsx");
                var work_sheet = workbook.Worksheets.First();

                PicklistXlsxFileCreator.Create(work_sheet, picklist);

                workbook.Save();
            }
        }
    }

    
}
