using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ListMakerOne;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    /*
            var reader = new SourceReader(source_sheet, parameters);
            var elements = SourceReader.GetElements();
            var buyers = SourceReader.GetBuyers();
            var amounts = SourceReader.GetAmounts();
            */

    public interface ISourceReader
    {
        IList<Element> GetElements();
        IList<string> GetBuyers();
        IList<ElementReservation> GetAmounts();
    }

    public class SourceReader : ISourceReader
    {
        private IXLWorksheet _work_sheet;
        private InputParameters _parameters;

        public SourceReader(IXLWorksheet work_sheet, InputParameters parameters)
        {
            _work_sheet = work_sheet;
            _parameters = parameters;
        }

        public IList<ElementReservation> GetAmounts() // ToDo test
        {
            var result = new List<ElementReservation>();

            var Element_Row_Span_Start = 0;
            var Element_Row_Span_End = 0;
            SettingsHelper.ReadSpan(_parameters.ElementRowSpan,
                out Element_Row_Span_Start, out Element_Row_Span_End);

            var Buyers_Column_Span_Start = "";
            var Buyers_Column_Span_End = "";
            SettingsHelper.ReadSpan(_parameters.BuyersColumnSpan,
                out Buyers_Column_Span_Start, out Buyers_Column_Span_End);

            for (int current_element_Row = Element_Row_Span_Start; 
                current_element_Row <= Element_Row_Span_End; 
                current_element_Row++)
            {
                var element_cell = _work_sheet.Cell(current_element_Row, _parameters.ElementIdColumn);
                var amount_cell = _work_sheet.Cell(current_element_Row, Buyers_Column_Span_Start);
                var buyer_cell = _work_sheet.Cell(_parameters.BuyersRow, Buyers_Column_Span_Start);

                while (true)
                {
                    var amount_raw = amount_cell.Value.ToString().Trim();
                    int amount = 0;

                    if (amount_raw != "" && int.TryParse(amount_raw, out amount))
                    {
                        if(amount > 0)
                        {
                            var elementid = _work_sheet.Cell(current_element_Row,
                                _parameters.ElementIdColumn).Value.ToString().Trim();
                            var buyer = buyer_cell.Value.ToString().Trim();

                            var reservation = new ElementReservation()
                            {
                                ElementID = elementid,
                                Receiver = buyer,
                                Amount = amount
                            };

                            result.Add(reservation);
                        }
                    }

                    if (amount_cell.Address.ColumnLetter == Buyers_Column_Span_End)
                        break;

                    amount_cell = amount_cell.CellRight();
                    buyer_cell = buyer_cell.CellRight();
                }

            }

            return result;
        }

        public IList<string> GetBuyers()
        {
            var result = new List<string>();

            var Buyers_Column_Span_Start = "";
            var Buyers_Column_Span_End = "";
            SettingsHelper.ReadSpan(_parameters.BuyersColumnSpan,
                out Buyers_Column_Span_Start, out Buyers_Column_Span_End);

            var current_col =  Buyers_Column_Span_Start;

            var buyer_cell = _work_sheet.Cell(_parameters.BuyersRow, Buyers_Column_Span_Start);

            while (true)
            {
                var buyer = buyer_cell.Value.ToString().Trim();
                result.Add(buyer);

                if (buyer_cell.Address.ColumnLetter == Buyers_Column_Span_End)
                    break;

                buyer_cell = buyer_cell.CellRight();
            }

            return result;
        }

        public IList<Element> GetElements()
        {
            var result = new List<Element>();

            var Element_Row_Span_Start = 0;
            var Element_Row_Span_End = 0;
            SettingsHelper.ReadSpan(_parameters.ElementRowSpan,
                out Element_Row_Span_Start, out Element_Row_Span_End);

            for (int i = Element_Row_Span_Start; i <= Element_Row_Span_End; i++)
            {
                var elementid = _work_sheet.Cell(i, _parameters.ElementIdColumn).Value.ToString().Trim();
                var description = _work_sheet.Cell(i, _parameters.BrickLinkDescriptionColumn).Value.ToString().Trim();
                var blid = _work_sheet.Cell(i, _parameters.BrickLinkIdColumn).Value.ToString().Trim();
                var blcolor = _work_sheet.Cell(i, _parameters.BrickLinkColorColumn).Value.ToString().Trim();

                result.Add(new Element()
                {
                    ElementID = elementid,
                    BricklinkDescription = description,
                    BricklinkId = blid,
                    BricklinkColor = blcolor,
                    MaterialColor = ""
                });
            }

            return result;
        }
    }
}
