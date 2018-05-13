using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace LugbulkListMaker
{
    public class MocGridWorker : IDataGridWorker
    {
        private MocGrid _grid;

        public MocGridWorker(MocGrid mocgrid)
        {
            this._grid = mocgrid;
        }

        public void ClearColumns()
        {
            //throw new NotImplementedException();
        }

        public void CreateColumns(int column_count)
        {
            //throw new NotImplementedException();
        }

        public void SetBackgroundColor(int row_start, int row_end, int column_start, int column_end, Color color)
        {
            //throw new NotImplementedException();
        }
    }
}
