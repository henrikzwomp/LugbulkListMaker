﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using ClosedXML.Excel;
using System.Windows.Controls;
using System.Windows.Media;

namespace LugbulkListMaker
{
    /*
     * ToDo
     * - Reload on Sheet selection change
     * - Test everything :)
     * */

    public class MainWindowLogic : ViewModelBase
    {
        IOutsideWindowHelper _outside_helper;
        string _selected_file_path;
        IXLWorkbook _workbook = null;
        IDataGridWorker _input_data_grid;

        const string _no_file_selected_text = "[None]";

        public MainWindowLogic(IOutsideWindowHelper outside_helper, IDataGridWorker input_data_grid)
        {
            _outside_helper = outside_helper;
            _input_data_grid = input_data_grid; 
        }

        private void SelectAInputFile()
        {
            _selected_file_path = "";
            _outside_helper.ShowLoadFileDialog(null, out _selected_file_path);

            UpdateSelectFileName();
            LoadSelectedWorkbook();
        }

        private void UpdateSelectFileName()
        {
            if(string.IsNullOrEmpty(_selected_file_path))
            {
                SelectFileName = _no_file_selected_text;
                return;
            }

            SelectFileName = _selected_file_path.Substring(_selected_file_path.LastIndexOf("\\")+1);
        }

        private void LoadSelectedWorkbook()
        {
            if (string.IsNullOrEmpty(_selected_file_path))
            {
                _workbook = null;
            }

            // ToDo check that file exists?

            _workbook = _outside_helper.GetXLWorkbook(_selected_file_path);

            UpdateSheetNames();
            UpdateGrid();
        }

        private void UpdateSheetNames()
        {
            SheetNames.Clear();

            if (_workbook == null)
            {
                SelectedSheetIndex = -1;
                IsFileLoaded = false;
                return;
            }
                

            foreach(var sheet in _workbook.Worksheets)
            {
                SheetNames.Add(sheet.Name);
            }

            IsFileLoaded = true;
            SelectedSheetIndex = 0;
        }

        private void UpdateGrid()
        {
            FileData.Clear();

            if (SelectedSheetIndex == -1)
            {
                _input_data_grid.ClearColumns();
                return;
            }

            var sheet = _workbook.Worksheet(SelectedSheetIndex+1);

            var x1 = sheet.LastCellUsed();
            var x2 = x1.Address;
            var x3 = x2.ColumnNumber;

            var cols = sheet.LastCellUsed().Address.ColumnNumber;
            var rows = sheet.LastCellUsed().Address.RowNumber;

            _input_data_grid.CreateColumns(cols);
            
            for (int i = 1; i <= rows; i++)
            {
                var values = new List<string>();
                values.Add(i.ToString());

                for (int j = 1; j <= cols; j++)
                {
                    values.Add(sheet.Cell(i, j).Value.ToString());
                }

                //_input_data_grid.Items.Add(values);
                FileData.Add(values);
            }
        }
        private Color ValidateSpanText(string span_text)
        {
            if(string.IsNullOrEmpty(span_text))
                return Colors.White;

            if(XLHelper.IsValidRangeAddress(span_text))
                return Colors.LightGreen;

            return Colors.LightPink;
        }

        #region PropertiesFields
        private string _selected_file_name = _no_file_selected_text;
        private ObservableCollection<string> _sheet_names = new ObservableCollection<string>();
        private int _selected_sheet_index = -1;
        private bool _selected_sheet_combobox_enable;
        private ObservableCollection<List<string>> _file_data = new ObservableCollection<List<string>>();
        string _element_id_span_text = "";
        string _buyers_names_span_text = "";
        string _bl_desc_span_text = "";
        string _bl_color_span_text = "";
        string _tlg_color_span_text = "";
        private SolidColorBrush _element_id_span_background = new SolidColorBrush(Colors.White);
        private SolidColorBrush _buyers_names_span_background = new SolidColorBrush(Colors.White);
        private SolidColorBrush _bl_desc_span_background = new SolidColorBrush(Colors.White);
        private SolidColorBrush _bl_color_span_background = new SolidColorBrush(Colors.White);
        private SolidColorBrush _tlg_color_span_background = new SolidColorBrush(Colors.White);
        #endregion

        #region Binded Properties

        public string SelectFileName
        { get
            {
                return _selected_file_name;
            }
            set
            {
                _selected_file_name = value;
                PropertyHasChanged("SelectFileName");
            }
        }

        public ObservableCollection<List<string>> FileData
        {
            get
            {
                return _file_data;
            }
        }

        public ObservableCollection<string> SheetNames
        {
            get
            {
                return _sheet_names;
            }
        }

        public int SelectedSheetIndex
        {
            get
            {
                return _selected_sheet_index;
            }
            set
            {
                _selected_sheet_index = value;
                UpdateGrid();
                PropertyHasChanged("SelectedSheetIndex");
            }
        }

        public bool IsFileLoaded
        {
            get
            {
                return _selected_sheet_combobox_enable;
            }
            set
            {
                _selected_sheet_combobox_enable = value;
                PropertyHasChanged("IsFileLoaded");
            }
        }

        public string ElementIdSpanText
        {
            get
            {
                return _element_id_span_text;
            }
            set
            {
                _element_id_span_text = value;
                ElementIdSpanBackground.Color = ValidateSpanText(_element_id_span_text);
                PropertyHasChanged("ElementIdSpanText");
                PropertyHasChanged("ElementIdSpanBackground");
            }
        }
        
        public string BuyersNamesSpanText
        {
            get
            {
                return _buyers_names_span_text;
            }
            set
            {
                _buyers_names_span_text = value;
                BuyersNamesSpanBackground.Color = ValidateSpanText(_buyers_names_span_text);
                PropertyHasChanged("BuyersNamesSpanText");
                PropertyHasChanged("BuyersNamesSpanBackground");
            }
        }
        
        public string BlDescSpanText
        {
            get
            {
                return _bl_desc_span_text;
            }
            set
            {
                _bl_desc_span_text = value;
                BlDescSpanBackground.Color = ValidateSpanText(_bl_desc_span_text);
                PropertyHasChanged("BlDescSpanText");
                PropertyHasChanged("BlDescSpanBackground");
            }
        }
        
        public string BlColorSpanText
        {
            get
            {
                return _bl_color_span_text;
            }
            set
            {
                _bl_color_span_text = value;
                BlColorSpanBackground.Color = ValidateSpanText(_bl_color_span_text);
                PropertyHasChanged("BlColorSpanText");
                PropertyHasChanged("BlColorSpanBackground");
            }
        }
        
        public string TlgColorSpanText
        {
            get
            {
                return _tlg_color_span_text;
            }
            set
            {
                _tlg_color_span_text = value;
                TlgColorSpanBackground.Color = ValidateSpanText(_tlg_color_span_text);
                PropertyHasChanged("TlgColorSpanText");
                PropertyHasChanged("TlgColorSpanBackground");
            }
        }

        public SolidColorBrush ElementIdSpanBackground
        {
            get
            {
                return _element_id_span_background;
            }
        }

        public SolidColorBrush BuyersNamesSpanBackground
        {
            get
            {
                return _buyers_names_span_background;
            }
        }


        public SolidColorBrush BlDescSpanBackground
        {
            get
            {
                return _bl_desc_span_background;
            }
        }


        public SolidColorBrush BlColorSpanBackground
        {
            get
            {
                return _bl_color_span_background;
            }
        }


        public SolidColorBrush TlgColorSpanBackground
        {
            get
            {
                return _tlg_color_span_background;
            }
        }

        #endregion

        #region Commands

        public DelegateCommand SelectInputFile
        {
            get
            {
                return new DelegateCommand(SelectAInputFile);
            }
        }
        #endregion

        
    }
}