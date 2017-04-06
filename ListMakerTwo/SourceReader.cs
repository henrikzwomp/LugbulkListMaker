using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        IList<LugBulkElement> GetElements();
        IList<LugBulkReceiver> GetBuyers();
        IList<LugBulkReservation> GetAmounts();
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

        public IList<LugBulkReservation> GetAmounts()
        {
            var result = new List<LugBulkReservation>();

            var buyers_list = GetBuyers();

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

                            var reservation = new LugBulkReservation()
                            {
                                ElementID = elementid,
                                Receiver = buyers_list.Where(x => x.Name == buyer).First(),
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

        public IList<LugBulkReceiver> GetBuyers()
        {
            var result = new List<LugBulkReceiver>();

            var Buyers_Column_Span_Start = "";
            var Buyers_Column_Span_End = "";
            SettingsHelper.ReadSpan(_parameters.BuyersColumnSpan,
                out Buyers_Column_Span_Start, out Buyers_Column_Span_End);

            var current_col =  Buyers_Column_Span_Start;

            var buyer_cell = _work_sheet.Cell(_parameters.BuyersRow, Buyers_Column_Span_Start);

            var buyer_id = 100;

            while (true)
            {
                var buyer = buyer_cell.Value.ToString().Trim();
                result.Add(new LugBulkReceiver() { Name = buyer, Id = buyer_id });

                if (buyer_cell.Address.ColumnLetter == Buyers_Column_Span_End)
                    break;

                buyer_cell = buyer_cell.CellRight();

                buyer_id++;
            }

            return result;
        }

        public IList<LugBulkElement> GetElements()
        {
            var result = new List<LugBulkElement>();

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

                result.Add(new LugBulkElement()
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
