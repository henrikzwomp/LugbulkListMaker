using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.IO;

namespace ListMakerTwo
{
    public class SqlFileCreator
    {
        public static IList<string> MakeFileForLugbulkDatabase(ISourceReader reader)
        {
            var lines = new List<string>();

            // Elements
            // [tblElements]: [ElementId], [BlId], [Description], [BLColor], [TlgColor], [TlgColorId]
            //      ,[Price], [SumQuantity], [Remainder]

            lines.Add("-- Elements --");
            var elements = reader.GetElements();
            foreach (var element in elements)
            {
                lines.Add(string.Format("INSERT INTO tblElements (ElementId, BlId, Description, BLColor) VALUES ({0}, '{1}', '{2}', '{3}')",
                    element.ElementID, element.BricklinkId, element.BricklinkDescription, element.BricklinkColor));
            }
            lines.Add("");

            // Buyers
            // [tblBuyers]: [Username], [MoneySum], [BrickAmount]
            lines.Add("-- Buyers --");
            var buyers = reader.GetBuyers();
            foreach (var buyer in buyers)
            {
                // [tblBuyers]: [Username], [MoneySum], [BrickAmount]
                lines.Add(string.Format("INSERT INTO tblBuyers (Username) VALUES ('{0}')", buyer.Name));
            }
            lines.Add("");

            // Amounts
            // [tblBuyersAmounts]: [Username], [ElementId], [Amount], [Difference]
            lines.Add("-- Amounts --");
            var amounts = reader.GetAmounts();
            foreach (var amount in amounts)
            {
                // [tblBuyers]: [Username], [MoneySum], [BrickAmount]
                lines.Add(string.Format("INSERT INTO tblBuyersAmounts (Username, ElementId, Amount) VALUES ('{0}', {1}, {2})",
                    amount.Receiver.Name, amount.ElementID, amount.Amount));
            }
            lines.Add("");

            //File.WriteAllLines("lugbulk_data.sql", lines);
            return lines;
        }
    }
}
