using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ListMakerTwo
{
    public interface ISourceReader
    {
        IList<LugBulkElement> GetElements();
        IList<LugBulkBuyer> GetBuyers();
        IList<LugBulkReservation> GetReservations();
    }
    
    public class SourceReader : ISourceReader
    {
        private IXLWorksheet _work_sheet;
        private InputParameters _parameters;

        IList<LugBulkReservation> _reservations;
        IList<LugBulkBuyer> _buyers;
        IList<LugBulkElement> _elements;

        public SourceReader(IXLWorksheet work_sheet, InputParameters parameters)
        {
            _work_sheet = work_sheet;
            _parameters = parameters;
        }

        public IList<LugBulkReservation> GetReservations()
        {
            if (_reservations != null)
                return _reservations;

            _reservations = new List<LugBulkReservation>();

            GetBuyers();
            GetElements();

            var reservation_pos_list = SourceReaderHelper.GetCrossRangePositions
                (_parameters.BuyersSpan, _parameters.ElementIdSpan);

            foreach(var reservation_pos in reservation_pos_list)
            {
                var reservation_amount = 0;
                
                // Skip if not a number
                if (!int.TryParse(_work_sheet.Cell(reservation_pos.Row, reservation_pos.Column)
                    .Value.ToString(), out reservation_amount))
                    continue;

                // Skip 0
                if (reservation_amount == 0)
                    continue;

                var element_id_pos = SourceReaderHelper.GetTitlePositionForValuePosition
                    (reservation_pos, _parameters.ElementIdSpan);
                var buyer_name_pos = SourceReaderHelper.GetTitlePositionForValuePosition
                    (reservation_pos, _parameters.BuyersSpan);

                var element_id = _work_sheet.Cell(element_id_pos.Row, element_id_pos.Column)
                    .Value.ToString().Trim();
                var buyer_name = _work_sheet.Cell(buyer_name_pos.Row, buyer_name_pos.Column)
                    .Value.ToString().Trim();

                var reservation = new LugBulkReservation()
                {
                    Element = _elements.Where(x => x.ElementID == element_id).First(),
                    Buyer = _buyers.Where(x => x.Name == buyer_name).First(),
                    Amount = reservation_amount,
                };

                _reservations.Add(reservation);
            }

            return _reservations;
        }

        public IList<LugBulkBuyer> GetBuyers()
        {
            if (_buyers != null)
                return _buyers;

            _buyers = new List<LugBulkBuyer>();

            var buyer_id = 100;

            var first_column = _parameters.BuyersSpan.FirstColumn().ColumnNumber();
            var last_column = _parameters.BuyersSpan.LastColumn().ColumnNumber();
            var first_row = _parameters.BuyersSpan.FirstRow().RowNumber();
            var last_row = _parameters.BuyersSpan.LastRow().RowNumber();

            CellPosition reservations_start_pos = null;
            CellPosition reservations_end_pos = null;

            SourceReaderHelper.GetCrossRangeStartEndPositions(
                _parameters.BuyersSpan, _parameters.ElementIdSpan, 
                out reservations_start_pos, out reservations_end_pos);

            for (int current_row = first_row; current_row <= last_row; current_row++)
            {
                for (int current_column = first_column; current_column <= last_column; current_column++)
                {
                    // Find a element amount

                    var reservation_values = SourceReaderHelper.GetValuesForTitlePosition(
                        new CellPosition() { Row = current_row, Column = current_column },
                        reservations_start_pos, reservations_end_pos, _work_sheet);

                    bool reserveration_found = false;
                    foreach(var value in reservation_values)
                    {
                        if (value == "")
                            continue;

                        int out_value = 0;
                        if(int.TryParse(value, out out_value))
                        {
                            if(out_value > 0)
                            {
                                reserveration_found = true;
                                break;
                            }
                        }
                    }

                    string buyer = _work_sheet.Cell(current_row, current_column).Value.ToString().Trim();

                    if (reserveration_found)
                    {
                        _buyers.Add(new LugBulkBuyer() { Name = buyer, Id = buyer_id });
                        buyer_id++;
                    }

                }
            }

            return _buyers;
        }

        public IList<LugBulkElement> GetElements()
        {
            if (_elements != null)
                return _elements;

            _elements = new List<LugBulkElement>();

            var element_id_values = ReadRangeValues(_parameters.ElementIdSpan);
            var bl_desc_values = ReadRangeValues(_parameters.BrickLinkDescriptionSpan);
            var bl_color_values = ReadRangeValues(_parameters.BrickLinkColorSpan);
            var bl_id_values = ReadRangeValues(_parameters.BrickLinkIdSpan);
            var tlg_color_values = ReadRangeValues(_parameters.TlgColorSpan);

            for(int i = 0; i < element_id_values.Count; i++)
            {
                _elements.Add(new LugBulkElement()
                {
                    ElementID = element_id_values[i],
                    BricklinkDescription = bl_desc_values[i],
                    BricklinkId = bl_id_values[i],
                    BricklinkColor = bl_color_values[i],
                    MaterialColor = tlg_color_values[i]
                });
            }

            return _elements;
        }

        private IList<string> ReadRangeValues(IXLRange range)
        {
            var result = new List<string>();

            var first_column = range.FirstColumn().ColumnNumber();
            var last_column = range.LastColumn().ColumnNumber();
            var first_row = range.FirstRow().RowNumber();
            var last_row = range.LastRow().RowNumber();

            for (int current_row = first_row; current_row <= last_row; current_row++)
            {
                for (int current_column = first_column; current_column <= last_column; current_column++)
                {
                    result.Add(_work_sheet.Cell(current_row, current_column).Value.ToString().Trim());
                }
            }

            return result;
        }
        
    }

    public class CellPosition
    {
        public int Row;
        public int Column;
    }
    
}

