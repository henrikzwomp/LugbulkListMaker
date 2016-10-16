using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using TeaboxDataFormat.IO;

namespace ListMakerOne
{
    public class Concept03 : Concept02
    {
        public static new void MainOne(string[] args)
        {
            // Load Element data 
            var elements = new List<Element>();
            var element_data_file = TeaboxDataFile.Open("Input2016\\tblElements.txt");
            foreach (var dataline in element_data_file.GetData())
            {
                elements.Add(new Element()
                {
                    ElementID = dataline["ElementId"],
                    BricklinkDescription = dataline["Description"],
                    BricklinkColor = dataline["BLColor"],
                    MaterialColor = dataline["TlgColor"]
                });
            }

            // Load Reservations list
            var reservations = new List<ElementReservation>();
            var reservations_data_file = TeaboxDataFile.Open("Input2016\\tblBuyersAmounts.txt");
            foreach (var dataline in reservations_data_file.GetData())
            {
                int amount = 0;
                if (!int.TryParse(dataline["Amount"], out amount))
                    amount = 0;

                reservations.Add(new ElementReservation()
                {
                    ElementID = dataline["ElementId"],
                    Receiver = dataline["Username"],
                    Amount = amount
                });
            }

            var picklists = ElementPicklistCreator.CreateLists(reservations, elements);

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
            File.WriteAllText("result.txt", output.ToString());
        }

    }
}
