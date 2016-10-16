using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerOne
{
    public class TemplateSettings
    {
        /* BaseTemplate02
        public int ElementDataColumn = 2;
        public int ElementDataFirstPageRowStart = 1;
        public int ElementDataSecondPageRowStart = 56;
        public int ReservationsFirstPageStartRow = 2;
        public int ReservationsFirstPageEndRow = 55;
        public int ReservationsReceiverCountColumn = 5;
        public int ReservationsReceiverColumn = 6;
        public int ReservationsAmountColumn = 7;
        public int ReservationsRecivedColumn = 8;
        public int ReservationsSecondPageEndRow = 102;
        public int ImportantInfoUpToRow = 10;
        public int SecondPageDescriptionFieldRow = 57;
        public int SecondPageDescriptionFieldCol = 2;
        public int SecondPageInstructionsRow = 65;
        public int SecondPageInstructionsCol = 1;
        */
        public int ElementDataColumn = 2;
        public int ElementDataFirstPageRowStart = 1;
        public int ElementDataSecondPageRowStart = 57;
        public int ReservationsFirstPageStartRow = 2;
        public int ReservationsFirstPageEndRow = 56;
        public int ReservationsReceiverCountColumn = 5;
        public int ReservationsReceiverColumn = 6;
        public int ReservationsAmountColumn = 7;
        public int ReservationsRecivedColumn = 8;
        public int ReservationsSecondPageEndRow = 102;
        public int ImportantInfoUpToRow = 10;
        public int SecondPageDescriptionFieldRow = 58;
        public int SecondPageDescriptionFieldCol = 2;
        public int SecondPageInstructionsRow = 66;
        public int SecondPageInstructionsCol = 1;
    }

    public class TemplateHandler
    {
        private TemplateSettings _settings;

        public TemplateHandler(TemplateSettings settings)
        {
            _settings = settings;
        }

        public void WriteInformationOnSecondPage(IXlsxFileContence file, ElementPicklist list)
        {
            file.SetCell(_settings.ElementDataSecondPageRowStart, _settings.ElementDataColumn, list.ElementID);
            file.SetCell(_settings.ElementDataSecondPageRowStart+1, _settings.ElementDataColumn, list.BricklinkDescription);
            file.SetCell(_settings.ElementDataSecondPageRowStart+4, _settings.ElementDataColumn, list.BricklinkColor);
            file.SetCell(_settings.ElementDataSecondPageRowStart+5, _settings.ElementDataColumn, list.MaterialColor);
        }

        public void WriteInformationOnFirstPage(IXlsxFileContence file, ElementPicklist list)
        {
            file.SetCell(_settings.ElementDataFirstPageRowStart, _settings.ElementDataColumn, list.ElementID);
            file.SetCell(_settings.ElementDataFirstPageRowStart+1, _settings.ElementDataColumn, list.BricklinkDescription);
            file.SetCell(_settings.ElementDataFirstPageRowStart+4, _settings.ElementDataColumn, list.BricklinkColor);
            file.SetCell(_settings.ElementDataFirstPageRowStart+5, _settings.ElementDataColumn, list.MaterialColor);
        }

        public void WriteReservations(IXlsxFileContence file, IList<ElementReservation> reservations)
        {
            int max_number_of_reservations_on_first_page = _settings.ReservationsFirstPageEndRow - _settings.ReservationsFirstPageStartRow + 1;
            int max_number_of_rows_on_first_page = _settings.ReservationsFirstPageEndRow;
            int max_number_of_rows = _settings.ReservationsSecondPageEndRow;

            int row = _settings.ReservationsFirstPageStartRow;

            foreach (var reservation in reservations)
            {
                file.SetCell(row, _settings.ReservationsReceiverColumn, reservation.Receiver);
                file.SetCell(row, _settings.ReservationsAmountColumn, reservation.Amount.ToString());

                if (row == max_number_of_rows_on_first_page && reservations.Count > max_number_of_reservations_on_first_page)
                    row++;

                row++;
            }

            while (row <= max_number_of_rows)
            {
                if(reservations.Count > max_number_of_reservations_on_first_page) // If "row" is pasted the first page
                {
                    if (row > (max_number_of_rows_on_first_page + _settings.ImportantInfoUpToRow)) 
                        file.DeleteRow(row);
                    else
                    {
                        file.DeleteCell(row, _settings.ReservationsReceiverCountColumn);
                        file.DeleteCell(row, _settings.ReservationsReceiverColumn);
                        file.DeleteCell(row, _settings.ReservationsAmountColumn);
                        file.DeleteCell(row, _settings.ReservationsRecivedColumn);
                    }
                }
                else
                {
                    if (row > _settings.ImportantInfoUpToRow)
                        file.DeleteRow(row);
                    else
                    {
                        file.DeleteCell(row, _settings.ReservationsReceiverCountColumn);
                        file.DeleteCell(row, _settings.ReservationsReceiverColumn);
                        file.DeleteCell(row, _settings.ReservationsAmountColumn);
                        file.DeleteCell(row, _settings.ReservationsRecivedColumn);
                    }
                }

                row++;
            }
        }

        public void RemoveMergedCellsOnSecondPage(IXlsxFileContence file)
        {
            file.RemoveMergeData(_settings.SecondPageDescriptionFieldRow, _settings.SecondPageDescriptionFieldCol);
            file.RemoveMergeData(_settings.SecondPageInstructionsRow, _settings.SecondPageInstructionsCol);
        }
    }
}
