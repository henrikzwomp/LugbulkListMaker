using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace LugbulkListMaker
{
    public class SheetData : IEnumerable<IEnumerable<SheetDataCell>>
    {
        List<IList<SheetDataCell>> _data;
        int current_row = -1;

        public SheetData()
        {
            _data = new List<IList<SheetDataCell>>();
        }

        public IEnumerator<IEnumerable<SheetDataCell>> GetEnumerator()
        {
            return _data.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _data.GetEnumerator();
        }

        internal void Add(string v)
        {
            if(current_row == -1)
                throw new Exception("Must call NewRow first"); // ToDo Test

            _data[current_row].Add(new SheetDataCell() { CellValue = v, BackgroundColor = Colors.White });
        }

        internal void NewRow()
        {
            _data.Add(new List<SheetDataCell>());
            current_row++;
        }

        public IList<IList<SheetDataCell>> ToList()
        {
            return _data;
        }
    }

    public class SheetDataCell : DependencyObject
    {
        public static readonly DependencyProperty CellValueProperty = DependencyProperty.Register("CellValue", typeof(string),
            typeof(SheetDataCell), new PropertyMetadata("")); // , new PropertyChangedCallback(ItemsSourceChanged)

        public static readonly DependencyProperty BackgroundColorProperty = DependencyProperty.Register("BackgroundColor", typeof(Color),
            typeof(SheetDataCell), new PropertyMetadata(Colors.White)); // , new PropertyChangedCallback(ItemsSourceChanged)


        public string CellValue
        {
            get { return (string)GetValue(CellValueProperty); }
            set { SetValue(CellValueProperty, value); } // Only called by code, never by WPF
        }

        public Color BackgroundColor
        {
            get { return (Color)GetValue(BackgroundColorProperty); }
            set { SetValue(BackgroundColorProperty, value); } // Only called by code, never by WPF
        }
    }
}
