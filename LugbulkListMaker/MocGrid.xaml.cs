﻿using System;
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
using System.Collections.ObjectModel;

using System.ComponentModel;
using System.Globalization;
using System.Diagnostics;

namespace LugbulkListMaker
{
    /// <summary>
    /// Interaction logic for MocGrid.xaml
    /// </summary>
    public partial class MocGrid : UserControl
    {
        const double CellWidth = 30;
        const double CellFontSize = 8;

        public MocGrid()
        {
            InitializeComponent();
            TestThingy.DataContext = this;

            _grid_cell_parent = new StackPanel();

            //TheScrollViewer.Content = _grid_cell_parent;

            /*var xxx = new SheetData();
            xxx.NewRow();
            xxx.Add("A1"); xxx.Add("A2"); xxx.Add("A3"); xxx.Add("A4");
            xxx.NewRow();
            xxx.Add("B1"); xxx.Add("B2"); xxx.Add("B3"); xxx.Add("B4");
            xxx.NewRow();
            xxx.Add("C1"); xxx.Add("C2"); xxx.Add("C3"); xxx.Add("C4");

            ItemsSource = xxx;*/
        }

        StackPanel _grid_cell_parent;
        bool _scale_view = false;
        int _max_columns = 0;

        private void Fill(IList<IList<SheetDataCell>> grid_data)
        {
            _grid_cell_parent.Children.Clear();

            _max_columns = 0;
            var row_titles = new StackPanel() { Orientation = Orientation.Horizontal };

            _grid_cell_parent.Children.Add(row_titles);

            for (int row_count = 0; row_count < grid_data.Count; row_count++)
            {
                var new_row = new StackPanel() { Orientation = Orientation.Horizontal };

                // Title Cell
                new_row.Children.Add(CreateTitleCell((row_count + 1).ToString(), new Thickness(1, 0, 1, 1)));

                for (int column_count = 0; column_count < grid_data[row_count].Count; column_count++)
                {
                    new_row.Children.Add(CreateValueCell(grid_data[row_count][column_count].CellValue));
                }

                if (_max_columns < grid_data[row_count].Count)
                    _max_columns = grid_data[row_count].Count;

                _grid_cell_parent.Children.Add(new_row);
            }

            // Add cells to Title row
            row_titles.Children.Add(CreateTitleCell("", new Thickness(1, 1, 1, 1)));
            for (int i = 1; i <= _max_columns; i++)
            {
                row_titles.Children.Add(CreateTitleCell(GetColumnLetterFromNumber(i), new Thickness(0, 1, 1, 1)));
            }

            // 
            for (int row_count = 1; row_count < grid_data.Count; row_count++)
            {
                var current_stack = (StackPanel)_grid_cell_parent.Children[row_count];

                if (current_stack.Children.Count <= _max_columns)
                {
                    for (int i = current_stack.Children.Count; i <= _max_columns; i++)
                    {
                        current_stack.Children.Add(CreateValueCell(""));
                    }

                }
            }
        }

        /*public void ColorIn(int start_row, int start_column, int end_row, int end_column, Brush brush)
        {
            if (start_row < 1)
                start_row = 1;
            if (start_column < 1)
                start_column = 1;
            if (end_row > _grid_cell_parent.Children.Count - 1)
                end_row = _grid_cell_parent.Children.Count - 1;
            if (end_column > _max_columns)
                end_column = _max_columns;


            for (int r = start_row; r <= end_row; r++)
            {
                for (int c = start_column; c <= end_column; c++)
                {
                    var row_stack_panel = (StackPanel)_grid_cell_parent.Children[r];
                    var cell_border = (Border)row_stack_panel.Children[c];
                    cell_border.Background = brush;
                }
            }
        }*/

        private static Border CreateValueCell(string text_contence)
        {
            var border = new Border() { BorderBrush = Brushes.Black };

            var text_item = new TextBlock()
            { Width = CellWidth, FontSize = CellFontSize };

            var thickness = new Thickness(0, 0, 1, 1);
            border.BorderThickness = thickness;

            text_item.Text = text_contence;

            border.Child = text_item;
            return border;
        }

        private static Border CreateTitleCell(string contence, Thickness thickness)
        {
            var title_border = new Border()
            {
                BorderThickness = thickness,
                BorderBrush = Brushes.Black,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            var title_text = new TextBlock()
            {
                Text = contence,
                Width = CellWidth,
                FontSize = CellFontSize,
                Background = Brushes.LightGray,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center, 
            };
            title_border.Child = title_text;
            return title_border;
        }

        private void ViewButton_Click(object sender, RoutedEventArgs e)
        {
            if (_scale_view)
            {
                //TheViewBox.Child = null;
                //TheScrollViewer.Content = _grid_cell_parent;
                _scale_view = false;
            }
            else
            {
                //TheScrollViewer.Content = null;
                //TheViewBox.Child = _grid_cell_parent;
                _scale_view = true;
            }

            /*foreach(var item in ItemsSource)
            {
                foreach(var item2 in item)
                {
                    item2.CellValue = "XXX!";
                    item2.BackgroundColor = Brushes.AliceBlue;
                }
            }*/
        }

        private static readonly string[] letters = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        private static string GetColumnLetterFromNumber(int column)
        {
            column--;
            if (column <= 25)
            {
                return letters[column];
            }
            var first_part = (column) / 26;
            var remainder = ((column) % 26) + 1;
            return GetColumnLetterFromNumber(first_part) + GetColumnLetterFromNumber(remainder);
        }

        public static readonly DependencyProperty ItemsSourceProperty =
      DependencyProperty.Register("ItemsSource", typeof(SheetData),
        typeof(MocGrid), new PropertyMetadata(new SheetData(), new PropertyChangedCallback(ItemsSourceChanged)));

        private static void ItemsSourceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            //var o = (MocGrid)d;
            //o.Fill(o.ItemsSource.ToList());
        }

        public SheetData ItemsSource
        {
            get { return (SheetData)GetValue(ItemsSourceProperty); }
            set { SetValue(ItemsSourceProperty,value); } // Only called by code, never by WPF
        }
    }

}
