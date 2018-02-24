using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using ClosedXML.Excel;
using System.Collections;
using System.Data;
using System.Windows.Media;

namespace LugbulkListMaker
{
    public interface IDataGridWorker
    {
        void ClearColumns();
        void CreateColumns(int column_count);
        void SetBackgroundColor(int row_start, int row_end, int column_start, int column_end, Color color);
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

        public void SetBackgroundColor(int row_start, int row_end, int column_start, int column_end, Color color)
        {
            int row_count = 0;

            _input_data_grid.UpdateLayout();

            for (int i = 0; i < _input_data_grid.Items.Count; i++)
            {
                var row = _input_data_grid.ItemContainerGenerator.ContainerFromIndex(i) as DataGridRow;

                if (null == row)
                    continue;

                int column_count = 0;

                foreach (DataGridColumn column in _input_data_grid.Columns)
                {
                    if (column.GetCellContent(row) is TextBlock)
                    {
                        TextBlock cellContent = column.GetCellContent(row) as TextBlock;

                        if (row_start <= row_count &&
                            row_end >= row_count &&
                            column_start <= column_count &&
                            column_end >= column_count)
                        {
                            _input_data_grid.ScrollIntoView(_input_data_grid.Items[row_count]);
                            cellContent.Background = new SolidColorBrush(color);
                        }
                            
                    }

                    column_count++;
                }

                row_count++;

            }
        }
    }
}
