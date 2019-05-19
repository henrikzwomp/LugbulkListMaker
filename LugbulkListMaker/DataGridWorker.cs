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

}
