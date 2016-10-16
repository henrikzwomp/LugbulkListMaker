using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerOne
{
    public class Concept02
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

            CreateXlsxFiles(picklists);
        }

        protected static void CreateXlsxFiles(IList<ElementPicklist> picklists)
        {
            foreach (var list in picklists)
            {
                var worker = new XlsxTemplateWorker("BaseTemplate03");

                var file = worker.GetXlsxFileContence();

                var template_settings = new TemplateSettings();
                var template_handler = new TemplateHandler(template_settings);

                template_handler.WriteInformationOnFirstPage(file, list);
                if (list.Reservations.Count >= template_settings.ReservationsFirstPageEndRow)
                    template_handler.WriteInformationOnSecondPage(file, list);
                else
                    template_handler.RemoveMergedCellsOnSecondPage(file);

                template_handler.WriteReservations(file, list.Reservations);

                worker.SaveToFile("Output", list.ElementID);
            }
        }
        // XlsxFileContence
        // TemplateHandler
        
    }
}
