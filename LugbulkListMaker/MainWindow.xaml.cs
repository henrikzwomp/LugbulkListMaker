using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Excel;

using System.Windows.Controls.Primitives;

namespace LugbulkListMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            DataContext = new MainWindowLogic(new OutsideWindowHelper(), new HighlightWorker(new TempObject()));

            InitializeComponent();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        class TempObject : IDataGridWorker
        {
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
    
}
