using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ListMakerOne;
using System.IO;

namespace ListMakerTwo
{
    public class Concept03Copy
    {
        public static void CreateListFiles(IList<ElementPicklist> picklists) // IList<ElementReservation> reservations, IList<Element> elements
        {
            // var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

            WriteToTextFile(picklists);

            CreateXlsxFiles(picklists);
        }

        private static void WriteToTextFile(IList<ElementPicklist> picklists)
        {
            // ...
            StringBuilder output = new StringBuilder();
            foreach (var picklist in picklists)
            {
                output.AppendLine("ElementID:\t" + picklist.ElementID);
                output.AppendLine("BL Description:\t" + picklist.BricklinkDescription);
                output.AppendLine(" ");
                output.AppendLine(" ");
                output.AppendLine("BL Color:\t" + picklist.BricklinkColor);
                output.AppendLine("TLG Color:\t" + picklist.MaterialColor);
                output.AppendLine(" ");

                int count = 1;
                foreach (var element_res in picklist.Reservations.OrderBy(x => x.Amount))
                {
                    output.AppendLine(count + "\t" + element_res.Receiver + "\t" + element_res.Amount);
                    count++;
                }

                output.AppendLine("\t");
                output.AppendLine("---------------------");
                output.AppendLine("\t");
            }


            // Profit!
            File.WriteAllText("TextCopy.txt", output.ToString());
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
    }
}
