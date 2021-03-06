﻿using System;
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
        private IXLRange _buyers_span;
        private IXLRange _element_id_span;
        private IXLRange _brickLink_description_span;
        private IXLRange _brickLink_id_span;
        private IXLRange _brickLink_color_span;
        private IXLRange _tlg_color_span;

        IList<LugBulkReservation> _reservations;
        IList<LugBulkBuyer> _buyers;
        IList<LugBulkElement> _elements;

        public SourceReader(IXLWorksheet work_sheet, InputParameters parameters)
        {
            _work_sheet = work_sheet;
            _buyers_span = work_sheet.Range(parameters.BuyersSpan);
            _element_id_span = work_sheet.Range(parameters.ElementIdSpan);
            _brickLink_description_span = work_sheet.Range(parameters.BrickLinkDescriptionSpan);
            _brickLink_id_span = work_sheet.Range(parameters.BrickLinkIdSpan);
            _brickLink_color_span = work_sheet.Range(parameters.BrickLinkColorSpan);
            _tlg_color_span = work_sheet.Range(parameters.TlgColorSpan);
        }

        public IList<LugBulkReservation> GetReservations()
        {
            if (_reservations != null)
                return _reservations;

            _reservations = new List<LugBulkReservation>();

            GetBuyers();
            GetElements();

            var reservation_pos_list = SourceReaderHelper.GetCrossRangePositions
                (_buyers_span, _element_id_span);

            foreach(var reservation_pos in reservation_pos_list)
            {
                var reservation_amount = 0;
                
                // Skip if not a number
                if (!int.TryParse(_work_sheet.Cell(reservation_pos.Row, reservation_pos.Column)
                    .Value.ToString(), out reservation_amount))
                    continue;

                // Skip 0 and less
                if (reservation_amount <= 0)
                    continue;

                var element_id_pos = SourceReaderHelper.GetTitlePositionForValuePosition
                    (reservation_pos, _element_id_span);
                var buyer_name_pos = SourceReaderHelper.GetTitlePositionForValuePosition
                    (reservation_pos, _buyers_span);

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

            var first_column = _buyers_span.FirstColumn().ColumnNumber();
            var last_column = _buyers_span.LastColumn().ColumnNumber();
            var first_row = _buyers_span.FirstRow().RowNumber();
            var last_row = _buyers_span.LastRow().RowNumber();

            for (int current_row = first_row; current_row <= last_row; current_row++)
            {
                for (int current_column = first_column; current_column <= last_column; current_column++)
                {
                    string buyer = _work_sheet.Cell(current_row, current_column).Value.ToString().Trim();

                    if (HasAReservation(current_row, current_column))
                    {
                        _buyers.Add(new LugBulkBuyer() { Name = buyer, Id = -1 });
                    }

                }
            }

            _buyers = _buyers.OrderBy(x => x.Name).ToList();

            var buyer_id = 100;

            foreach(var buyer in _buyers)
            {
                buyer.Id = buyer_id;
                buyer_id++;
            }

            return _buyers;
        }

        private CellPosition _reservations_start_pos = null;
        private CellPosition _reservations_end_pos = null;

        private CellPosition ReservationsStartPos
        {
            get
            {
                if(_reservations_start_pos == null)
                {
                    // ToDo duplicate code
                    SourceReaderHelper.GetCrossRangeStartEndPositions(
                        _buyers_span, _element_id_span,
                        out _reservations_start_pos, out _reservations_end_pos);
                }

                return _reservations_start_pos;
            }
        }

        private CellPosition ReservationsEndPos
        {
            get
            {
                if (_reservations_end_pos == null)
                {
                    // ToDo duplicate code
                    SourceReaderHelper.GetCrossRangeStartEndPositions(
                        _buyers_span, _element_id_span,
                        out _reservations_start_pos, out _reservations_end_pos);
                }

                return _reservations_end_pos;
            }
        }

        private bool HasAReservation(int current_row, int current_column)
        {
            var reservation_values = SourceReaderHelper.GetValuesForTitlePosition(
                new CellPosition() { Row = current_row, Column = current_column },
                ReservationsStartPos, ReservationsEndPos, _work_sheet);

            bool reserveration_found = false;
            foreach (var value in reservation_values)
            {
                if (value == "")
                    continue;

                int out_value = 0;
                if (int.TryParse(value, out out_value))
                {
                    if (out_value > 0)
                    {
                        reserveration_found = true;
                        break;
                    }
                }
            }

            return reserveration_found;
        }

        public IList<LugBulkElement> GetElements()
        {
            if (_elements != null)
                return _elements;

            _elements = new List<LugBulkElement>();

            var bl_desc_values = ReadRangeValues(_brickLink_description_span);
            var bl_color_values = ReadRangeValues(_brickLink_color_span);
            var bl_id_values = ReadRangeValues(_brickLink_id_span);
            var tlg_color_values = ReadRangeValues(_tlg_color_span);

            var first_column = _element_id_span.FirstColumn().ColumnNumber();
            var last_column = _element_id_span.LastColumn().ColumnNumber();
            var first_row = _element_id_span.FirstRow().RowNumber();
            var last_row = _element_id_span.LastRow().RowNumber();

            var element_counter = 0;

            for (int current_row = first_row; current_row <= last_row; current_row++)
            {
                for (int current_column = first_column; current_column <= last_column; current_column++)
                {
                    if (HasAReservation(current_row, current_column))
                    {
                        var element_id = _work_sheet.Cell(current_row, current_column).Value.ToString().Trim();

                        _elements.Add(new LugBulkElement()
                        {
                            ElementID = element_id,
                            BricklinkDescription = bl_desc_values[element_counter],
                            BricklinkId = bl_id_values[element_counter],
                            BricklinkColor = bl_color_values[element_counter],
                            MaterialColor = tlg_color_values[element_counter]
                        });
                    }

                    element_counter++;
                }
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

