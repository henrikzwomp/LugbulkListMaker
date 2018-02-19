using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using ClosedXML.Excel;

namespace LugbulkListMaker
{
    public interface IDataGridWorker
    {
        void ClearColumns();
        void CreateColumns(int column_count);
    }

    public class DataGridWorker : IDataGridWorker
    {
        private DataGrid _input_data_grid;

        public DataGridWorker(DataGrid input_data_grid)
        {
            _input_data_grid = input_data_grid;
        }

        public void ClearColumns()
        {
            _input_data_grid.Columns.Clear();
        }

        public void CreateColumns(int column_count)
        {
            _input_data_grid.Columns.Clear();

            for (int i = 1; i <= column_count; i++)
            {
                _input_data_grid.Columns.Add(new DataGridTextColumn() { Header = XLHelper.GetColumnLetterFromNumber(i), Binding = new System.Windows.Data.Binding("[" + i + "]") });
            }
            
        }
    }
}
