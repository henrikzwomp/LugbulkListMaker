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
            string INPUT_FILE_NAME = "LUGBULK 2017 beställning färdig.xlsx";
            string WORKSHEET_NAME = "Beställning sammanställd";
            /*
            _Requiered_
            Element_Row_Span = 2:86
            Buyers_Row = 87
            Buyers_Column_Span = K:FW
            Element_Id_Column = D

            _Optional_
            BrickLink_Description_Column = B
            BrickLink_Id_Column = C
            BrickLink_Color_Column = E
            Tlg_Color = 
            */

            var workbook = new XLWorkbook(INPUT_FILE_NAME);
            var sheet = workbook.Worksheets.First(x => x.Name == WORKSHEET_NAME);

            //PrintElementIds(sheet);
            //PrintBuyers(sheet);
            ToSqls(sheet);
        }

        private static void PrintElementIds(IXLWorksheet sheet)
        {
            for (int i = 2; i <= 86; i++)
            {
                var cell_name = "D" + i;

                var cell = sheet.Cell(i, "D");



                if (cell == null)
                    Console.WriteLine(cell_name + ": Null");
                else
                    Console.WriteLine(cell_name + ": " + cell.Value);
            }
        }

        private static void PrintBuyers(IXLWorksheet sheet)
        {
            var cell = sheet.Cell("K87");

            while (true)
            {
                Console.WriteLine(cell.Address.ToStringRelative() + ": " + cell.Value);
                
                if (cell.Address.ToStringRelative() == "FW87")
                    break;

                cell = cell.CellRight();
            }
        }


        private static void ToSqls(IXLWorksheet sheet)
        {
            var Element_Row_Span_Start = 0;
            var Element_Row_Span_End = 0;
            SettingsHelper.ReadSpan("2:86", out Element_Row_Span_Start, out Element_Row_Span_End);

            var Buyers_Column_Span_Start = "";
            var Buyers_Column_Span_End = "";
            SettingsHelper.ReadSpan("K:FW", out Buyers_Column_Span_Start, out Buyers_Column_Span_End);

            var Buyers_Row = "87";
            var Element_Id_Column = "D";
            var BrickLink_Description_Column = "B";
            var BrickLink_Id_Column = "C";
            var BrickLink_Color_Column = "E";

            var lines = new List<string>();

            // Elements
            // [tblElements]: [ElementId], [BlId], [Description], [BLColor], [TlgColor], [TlgColorId]
            //      ,[Price], [SumQuantity], [Remainder]

            lines.Add("-- Elements --");
            for (int i = Element_Row_Span_Start; i <= Element_Row_Span_End; i++)
            {
                var elementid_cell = sheet.Cell(i, Element_Id_Column).Value.ToString().Trim();
                var description_cell = sheet.Cell(i, BrickLink_Description_Column).Value.ToString().Trim();
                var blid_cell = sheet.Cell(i, BrickLink_Id_Column).Value.ToString().Trim();
                var blcolor_cell = sheet.Cell(i, BrickLink_Color_Column).Value.ToString().Trim();

                lines.Add(string.Format("INSERT INTO tblElements (ElementId, BlId, Description, BLColor) VALUES ({0}, '{1}', '{2}', '{3}')", 
                    elementid_cell, blid_cell, description_cell, blcolor_cell));
            }
            lines.Add("");

            // Buyers
            // [tblBuyers]: [Username], [MoneySum], [BrickAmount]
            lines.Add("-- Buyers --");
            var current_col = Buyers_Column_Span_Start;

            var buyer_cell = sheet.Cell(Buyers_Column_Span_Start + Buyers_Row);

            while (true)
            {
                // [tblBuyers]: [Username], [MoneySum], [BrickAmount]

                var buyer = buyer_cell.Value.ToString().Trim();

                lines.Add(string.Format("INSERT INTO tblBuyers (Username) VALUES ('{0}')", buyer));

                if (buyer_cell.Address.ColumnLetter == Buyers_Column_Span_End)
                    break;

                buyer_cell = buyer_cell.CellRight();
            }
            lines.Add("");

            // Amounts
            // [tblBuyersAmounts]: [Username], [ElementId], [Amount], [Difference]
            lines.Add("-- Amounts --");

            for (int Element_Row = Element_Row_Span_Start; Element_Row <= Element_Row_Span_End; Element_Row++)
            {
                var element_cell = sheet.Cell(Element_Id_Column + Element_Row);
                var amount_cell = sheet.Cell(Buyers_Column_Span_Start + Element_Row);

                buyer_cell = sheet.Cell(Buyers_Column_Span_Start + Buyers_Row);

                while (true)
                {
                    // [tblBuyers]: [Username], [MoneySum], [BrickAmount]

                    var amount = amount_cell.Value.ToString().Trim();

                    if(amount != "")
                    {
                        var elementid = sheet.Cell(Element_Row, Element_Id_Column).Value.ToString().Trim();
                        var buyer = buyer_cell.Value.ToString().Trim();

                        lines.Add(string.Format("INSERT INTO tblBuyersAmounts (Username, ElementId, Amount) VALUES ('{0}', {1}, {2})",
                            buyer, elementid, amount));
                    }

                    if (amount_cell.Address.ColumnLetter == Buyers_Column_Span_End)
                        break;

                    amount_cell = amount_cell.CellRight();
                    buyer_cell = buyer_cell.CellRight();
                }

            }
            lines.Add("");


            lines.Add("");

            File.WriteAllLines("sqls.txt", lines);
        }

        
    }
}
