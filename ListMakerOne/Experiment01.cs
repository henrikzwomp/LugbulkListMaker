using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace ListMakerOne
{
    class Experiment01
    {
        public static void Go()
        {
            string output_file_name = "result.xlsx";
            var dir_info = new DirectoryInfo("BaseTemplate");

            if (File.Exists(output_file_name))
                File.Delete(output_file_name);

            var xdoc = XDocument.Load(dir_info.Name + "\\" + "xl/worksheets/sheet1.xml");
            
            var sheet_data = xdoc.Elements().First()
                .Elements().Where(x => x.Name.LocalName == "sheetData").First();

            XNamespace ns = xdoc.Root.Name.Namespace;

            var row1 = new XElement(ns + "row");
            row1.Add(new XAttribute("r", "1"));
            row1.Add(NewColumn(ns, "A1", "2", "s", 0));
            row1.Add(NewColumn(ns, "B1", "3", "s", 10));
            row1.Add(NewColumn(ns, "C1", "4"));
            row1.Add(NewColumn(ns, "D1", "5"));
            row1.Add(NewColumn(ns, "E1", "8", "s", 1));
            row1.Add(NewColumn(ns, "F1", "9", "s", 14));
            row1.Add(NewColumn(ns, "G1", "10", "s", 2));
            row1.Add(NewColumn(ns, "H1", "10", "s", 3));
            sheet_data.Add(row1);

            using (ZipArchive archive = ZipFile.Open(output_file_name, ZipArchiveMode.Create))
            {
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "[Content_Types].xml", "[Content_Types].xml");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "_rels/.rels", "_rels/.rels");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/workbook.xml", "xl/workbook.xml");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/theme/theme1.xml", "xl/theme/theme1.xml");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/worksheets/_rels/sheet1.xml.rels", "xl/worksheets/_rels/sheet1.xml.rels");

                //archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml");
                var entry = archive.CreateEntry("xl/worksheets/sheet1.xml");
                using (Stream stream = entry.Open())
                {
                    xdoc.Save(stream);
                }

                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/styles.xml", "xl/styles.xml");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/sharedStrings.xml", "xl/sharedStrings.xml");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "docProps/core.xml", "docProps/core.xml");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/printerSettings/printerSettings1.bin", "xl/printerSettings/printerSettings1.bin");
                archive.CreateEntryFromFile(dir_info.Name + "\\" + "docProps/app.xml", "docProps/app.xml");
                //archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/calcChain.xml", "xl/calcChain.xml");
            }
        }

        private static XElement NewColumn(XNamespace ns, string r, string s, string t, int v)
        {
            var obj = new XElement(ns + "c");
            obj.Add(new XAttribute("r", r));
            obj.Add(new XAttribute("s", s));
            obj.Add(new XAttribute("t", t));
            obj.Add(NewSharedString(ns, v));
            return obj;
        }

        private static XElement NewColumn(XNamespace ns, string r, string s)
        {
            var obj = new XElement(ns + "c");
            obj.Add(new XAttribute("r", r));
            obj.Add(new XAttribute("s", s));
            return obj;
        }

        private static XElement NewSharedString(XNamespace ns, int v)
        {
            var obj = new XElement(ns + "v");
            obj.Value = v.ToString();
            return obj;
        }

    }

    
}
