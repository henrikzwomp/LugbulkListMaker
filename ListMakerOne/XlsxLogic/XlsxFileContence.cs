using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ListMakerOne
{
    public enum XlsxFileStyle
    {
        Bold_GreyBackground,
        GreyBackground,
        Bold,
        None
    }

    public interface IXlsxFileContence
    {
        void SetCell(int row_index, int column_index, string text);
        void AddCell(int row_index, int column_index, string text, XlsxFileStyle style);
        void AddCell(int row, int column, string text, XlsxFileStyle style, int merge_rows, int merge_cols);
        void DeleteCell(int row_index, int column_index);
        void DeleteRow(int row_index);
        void RemoveMergeData(int row_index, int column_index);
    }

    public class XlsxFileContence : IXlsxFileContence
    {
        private XDocument _xdoc;
        private XElement _sheet_data;
        private XElement _merge_cells;
        private XNamespace _ns;

        public XlsxFileContence(XDocument xdoc)
        {
            _xdoc = xdoc;
            _sheet_data = xdoc.Elements().First()
                .Elements().Where(x => x.Name.LocalName == "sheetData").First();
            _merge_cells = xdoc.Elements().First()
                .Elements().Where(x => x.Name.LocalName == "mergeCells").First();
            _ns = xdoc.Root.Name.Namespace;
        }

        public void AddCell(int row_index, int column_index, string text, XlsxFileStyle style)
        {
            var column_letter = ToColumnLetter(column_index);
            var style_id = ToStyleId(style);

            var row = new XElement(_ns + "row");
            row.Add(new XAttribute("r", row_index));

            row.Add(NewColumn(column_letter + row_index, style_id, text));

            // <row r="3" >
            // <c r="F3" s="12" t="str"><v>#Reciever2#</v></c>

            _sheet_data.Add(row);
        }

        public void AddCell(int row, int column, string text, XlsxFileStyle style, int merge_rows, int merge_cols)
        {
            AddCell(row, column, text, style);
            SetMerge(row, column, merge_rows, merge_cols);
        }

        public void SetCell(int row_index, int column_index, string text)
        {
            var column_lettet = ToColumnLetter(column_index);
            XElement row = GetRow(row_index);
            XElement cell = GetCell(column_lettet + row_index, row);

            var v = cell.Elements().Where(x => x.Name.LocalName == "v").First();
            v.Value = text;

            var t = cell.Attributes().Where(x => x.Name.LocalName == "t").First();
            t.Value = "str";
        }

        public void DeleteCell(int row_index, int column_index)
        {
            var column_lettet = ToColumnLetter(column_index);

            XElement row = GetRow(row_index);

            GetCell(column_lettet + row_index, row).Remove();

            if (!row.Elements().Any())
                row.Remove();
        }

        public void DeleteRow(int row_index)
        {
            XElement row = GetRow(row_index);
            if(row != null)
                row.Remove();
        }

        public void RemoveMergeData(int row_index, int column_index)
        {
            var count_attribute = _merge_cells.Attributes().Where(x => x.Name == "count").First();
            int count = int.Parse(count_attribute.Value);
            count_attribute.Value = (--count).ToString();

            var start_column_letter = ToColumnLetter(column_index);

            string reftext_start = start_column_letter + (row_index) + ":";

            var merge_cell = _merge_cells.Elements().Where(x => x.Name.LocalName == "mergeCell" && x.Attributes().Where(y => y.Name.LocalName == "ref" && y.Value.StartsWith(reftext_start)).Any()).First();
            merge_cell.Remove();
        }

        private int ToStyleId(XlsxFileStyle style)
        {
            if (style == XlsxFileStyle.None) return -1;
            if (style == XlsxFileStyle.Bold) return 0;
            if (style == XlsxFileStyle.Bold_GreyBackground) return 0;
            if (style == XlsxFileStyle.GreyBackground) return 10;
            throw new Exception("No id for given XlsxFileStyle found");
        }

        private string ToColumnLetter(int column_index)
        {
            if (column_index == 1) return "A";
            if (column_index == 2) return "B";
            if (column_index == 3) return "C";
            if (column_index == 4) return "D";
            if (column_index == 5) return "E";
            if (column_index == 6) return "F";
            if (column_index == 7) return "G";
            if (column_index == 8) return "H";
            if (column_index == 9) return "I";
            if (column_index == 10) return "J";
            throw new Exception("No letter for given column_index found");
        }

        private void SetMerge(int start_row, int start_column, int merge_rows, int merge_cols)
        {
            var count_attribute = _merge_cells.Attributes().Where(x => x.Name == "count").First();
            int count = int.Parse(count_attribute.Value);
            count_attribute.Value = (++count).ToString();


            var start_column_letter = ToColumnLetter(start_column);

            var end_column_letter = ToColumnLetter(start_column + merge_cols - 1);

            string reftext = start_column_letter + (start_row) + ":" + end_column_letter + (start_row + merge_rows - 1);

            var merge_cell = new XElement(_ns + "mergeCell");
            merge_cell.Add(new XAttribute("ref", reftext));

            _merge_cells.Add(merge_cell);
        }

        private XElement NewColumn(string cell_id, int style_id, string text)
        {
            var obj_c = new XElement(_ns + "c");
            obj_c.Add(new XAttribute("r", cell_id));
            if(style_id > -1)
                obj_c.Add(new XAttribute("s", style_id));
            obj_c.Add(new XAttribute("t", "str"));

            var obj_v = new XElement(_ns + "v");
            obj_v.Value = text;
            obj_c.Add(obj_v);

            return obj_c;
        }

        private static XElement GetCell(string cell_id, XElement row)
        {
            return row.Elements().Where(x => x.Name.LocalName == "c" && x.Attributes().Where(y => y.Name.LocalName == "r" && y.Value == cell_id).Any()).First();
        }

        private XElement GetRow(int row_index)
        {
            var rows = _sheet_data.Elements().Where(x => x.Name.LocalName == "row" && x.Attributes().Where(y => y.Name.LocalName == "r" && y.Value == row_index.ToString()).Any());

            if (rows.Any())
                return rows.First();

            return null;
        }
    }
}
