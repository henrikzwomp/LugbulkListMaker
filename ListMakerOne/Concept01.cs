using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ListMakerOne
{
    class Concept01
    {
        public static void MainOne(string[] args)
        {
            // "Load" Reservations list
            var reservations = new List<ElementReservation>();
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10002", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Teabox", ElementID = "10003", Amount = 50 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10001", Amount = 200 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10002", Amount = 1000 });
            reservations.Add(new ElementReservation() { Receiver = "Simpson", ElementID = "10003", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10001", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10002", Amount = 100 });
            reservations.Add(new ElementReservation() { Receiver = "Alice", ElementID = "10003", Amount = 200 });

            // "Load" Element data 
            var elements = new List<Element>();
            elements.Add(new Element() { ElementID = "10001", BricklinkDescription = "Plant", BricklinkColor = "Green", MaterialColor = "Dark Green" });
            elements.Add(new Element() { ElementID = "10002", BricklinkDescription = "Bone", BricklinkColor = "White", MaterialColor = "White" });
            elements.Add(new Element() { ElementID = "10003", BricklinkDescription = "Brick 1 x 2", BricklinkColor = "Dark Red", MaterialColor = "New Dark Red" });

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            WriteToTextFile(picklists);

            CreateXlsxFiles(picklists);
        }

        private static void CreateXlsxFiles(IList<ElementPicklist> picklists)
        {
            foreach (var list in picklists)
            {
                var worker = new XlsxTemplateWorker("BaseTemplate01");

                var file = worker.GetXlsxFileContence();

                XlsxFileWriter.WriteInformationOnFirstPage(file, list);
                //if (list.Reservations.Count > XlsxFileWriter.FirstPageRowBreak)
                //    XlsxFileWriter.WriteInformationOnSecondPage(file, list); ToDo

                //XlsxFileWriter.WriteReservations(file, list.Reservations); ToDo

                worker.SaveToFile("Output", list.ElementID);
            }
        }

        class XlsxFileWriter
        {
            public const int FirstPageRowBreak = 54; // ??? ToDo check this

            public static void WriteReservations(IXlsxFileContence file, IList<ElementReservation> reservations)
            {
                throw new NotImplementedException();
            }

            public static void WriteInformationOnSecondPage(IXlsxFileContence file, ElementPicklist list)
            {
                throw new NotImplementedException();
            }

            public static void WriteInformationOnFirstPage(IXlsxFileContence file, ElementPicklist list)
            {
                string BigAnvisningarString = "";

                file.AddCell(1, 1, "ElementID:", XlsxFileStyle.Bold_GreyBackground);
                file.AddCell(1, 2, list.ElementID, XlsxFileStyle.GreyBackground);
                file.AddCell(1, 3, "", XlsxFileStyle.GreyBackground);

                file.AddCell(2, 1, "BL description:", XlsxFileStyle.Bold_GreyBackground);
                file.AddCell(2, 2, list.BricklinkDescription, XlsxFileStyle.GreyBackground, 3, 2);
                //file.SetCell(2, 3, "", XlsxFileStyle.GreyBackground, 2, 2);

                file.AddCell(3, 1, "", XlsxFileStyle.GreyBackground);
                file.AddCell(3, 2, "", XlsxFileStyle.GreyBackground);
                file.AddCell(3, 3, "", XlsxFileStyle.GreyBackground);
                file.AddCell(4, 1, "", XlsxFileStyle.GreyBackground);
                file.AddCell(4, 2, "", XlsxFileStyle.GreyBackground);
                file.AddCell(4, 3, "", XlsxFileStyle.GreyBackground);

                file.AddCell(5, 1, "BL Color:", XlsxFileStyle.Bold_GreyBackground);
                file.AddCell(5, 2, list.BricklinkColor, XlsxFileStyle.GreyBackground);
                file.AddCell(5, 3, "", XlsxFileStyle.GreyBackground);

                file.AddCell(6, 1, "TLG Color:", XlsxFileStyle.Bold_GreyBackground);
                file.AddCell(6, 2, list.MaterialColor, XlsxFileStyle.GreyBackground);
                file.AddCell(6, 3, "", XlsxFileStyle.GreyBackground);

                file.AddCell(9, 1, "Anvisningar", XlsxFileStyle.Bold);
                file.AddCell(10, 1, BigAnvisningarString, XlsxFileStyle.None, 29, 3);
            }
        }

        private static void WriteToTextFile(IList<ElementPicklist> picklists)
        {
            // ...
            StringBuilder output = new StringBuilder();
            foreach (var picklist in picklists)
            {
                output.AppendLine("ElementID: " + picklist.ElementID);
                output.AppendLine("BL Description: " + picklist.BricklinkDescription);
                output.AppendLine("BL Color: " + picklist.BricklinkColor);
                output.AppendLine("TLG Color: " + picklist.MaterialColor);
                output.AppendLine("");

                foreach (var element_res in picklist.Reservations.OrderBy(x => x.Amount))
                {
                    output.AppendLine(element_res.Receiver + "\t" + element_res.Amount);
                }

                output.AppendLine("");
                output.AppendLine("----------");
                output.AppendLine("");
            }


            // Profit!
            File.WriteAllText("result.txt", output.ToString());
        }
    }
}
