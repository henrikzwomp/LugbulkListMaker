using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace LugbulkListMaker
{
    public interface IHighlightWorker
    {
        void ClearHighlightColor(Color highlight);
        void SetOrUpdateHighlightColor(Color highlight, int row_start, int row_end, int column_start, int column_end);
    }

    // ToDO Test
    public class HighlightWorker : IHighlightWorker
    {
        private IDataGridWorker _input_data_grid;
        private Dictionary<Color, HighlightSpan> _highlights;

        public HighlightWorker(IDataGridWorker input_data_grid)
        {
            _input_data_grid = input_data_grid;
            _highlights = new Dictionary<Color, HighlightSpan>();
        }

        public void ClearHighlightColor(Color highlight)
        {
            RemoveHighlightColor(highlight);
            _highlights.Remove(highlight);
            ReHighlightCells();
        }

        public void SetOrUpdateHighlightColor(Color highlight, int row_start, int row_end, 
            int column_start, int column_end)
        {
            if(_highlights.ContainsKey(highlight))
            {
                RemoveHighlightColor(highlight);

                _highlights[highlight].RowStart = row_start;
                _highlights[highlight].RowEnd = row_end;
                _highlights[highlight].ColumnStart = column_start;
                _highlights[highlight].ColumnEnd = column_end;
            }
            else
            {
                _highlights.Add(highlight, new HighlightSpan() {
                    RowStart = row_start, 
                    RowEnd = row_end,
                    ColumnStart = column_start,
                    ColumnEnd = column_end,
                });
            }

            ReHighlightCells();
        }

        private void ReHighlightCells()
        {
            foreach(var item in _highlights)
            {
                var highlight_color = item.Key;
                var span = item.Value;

                _input_data_grid.SetBackgroundColor(span.RowStart-1, span.RowEnd-1,
                    span.ColumnStart-1, span.ColumnEnd-1, highlight_color);
            }

        }

        private void RemoveHighlightColor(Color highlight)
        {
            if(_highlights.ContainsKey(highlight))
            {
                var span = _highlights[highlight];

                _input_data_grid.SetBackgroundColor(span.RowStart-1, span.RowEnd-1,
                    span.ColumnStart-1, span.ColumnEnd-1, Colors.White);
            }
        }

        private class HighlightSpan
        {
            public int RowStart;
            public int RowEnd;
            public int ColumnStart;
            public int ColumnEnd;
        }
    }
}
