using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;

namespace LugbulkListMaker
{
    public interface IOutsideWindowHelper
    {
        bool ShowLoadFileDialog(string initial_directory, out string filename);
        IXLWorkbook GetXLWorkbook(string _selected_file_path);
    }

    public class OutsideWindowHelper : IOutsideWindowHelper
    {
        public bool ShowLoadFileDialog(string initial_directory, out string filename)
        {
            filename = null;

            if (string.IsNullOrEmpty(initial_directory) || !Directory.Exists(initial_directory))
                initial_directory = AppDomain.CurrentDomain.BaseDirectory;

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlmx;.xls";
            dlg.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls";
            dlg.InitialDirectory = initial_directory;

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                filename = dlg.FileName;
                return true;
            }

            return false;
        }

        public IXLWorkbook GetXLWorkbook(string file_path)
        {
            return new XLWorkbook(file_path);
        }

        
    }
}
