using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.IO.Compression;

namespace ListMakerOne
{
    public class XlsxTemplateWorker
    {
        private string _template_path;
        private XDocument _xdoc;

        public XlsxTemplateWorker(string template_path)
        {
            _template_path = template_path;

            if (!_template_path.EndsWith("\\"))
                _template_path += "\\";
        }

        public IXlsxFileContence GetXlsxFileContence()
        {
            _xdoc = XDocument.Load(_template_path + "xl/worksheets/sheet1.xml");
            return new XlsxFileContence(_xdoc);
        }

        public void SaveToFile(string output_folder, string file_name)
        {
            if (!Directory.Exists(output_folder))
                Directory.CreateDirectory(output_folder);

            if (!output_folder.EndsWith("\\"))
                output_folder += "\\";

            var output_file_path = output_folder + file_name + ".xlsx";

            if (File.Exists(output_file_path))
                File.Delete(output_file_path);

            using (ZipArchive archive = ZipFile.Open(output_file_path, ZipArchiveMode.Create))
            {
                archive.CreateEntryFromFile(_template_path + "[Content_Types].xml", "[Content_Types].xml");
                archive.CreateEntryFromFile(_template_path + "_rels/.rels", "_rels/.rels");
                archive.CreateEntryFromFile(_template_path + "xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels");
                archive.CreateEntryFromFile(_template_path + "xl/workbook.xml", "xl/workbook.xml");
                archive.CreateEntryFromFile(_template_path + "xl/theme/theme1.xml", "xl/theme/theme1.xml");
                archive.CreateEntryFromFile(_template_path + "xl/worksheets/_rels/sheet1.xml.rels", "xl/worksheets/_rels/sheet1.xml.rels");

                //archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml");
                var entry = archive.CreateEntry("xl/worksheets/sheet1.xml");
                using (Stream stream = entry.Open())
                {
                    _xdoc.Save(stream);
                }

                archive.CreateEntryFromFile(_template_path + "xl/styles.xml", "xl/styles.xml");
                archive.CreateEntryFromFile(_template_path + "xl/sharedStrings.xml", "xl/sharedStrings.xml");
                archive.CreateEntryFromFile(_template_path + "docProps/core.xml", "docProps/core.xml");
                archive.CreateEntryFromFile(_template_path + "xl/printerSettings/printerSettings1.bin", "xl/printerSettings/printerSettings1.bin");
                archive.CreateEntryFromFile(_template_path + "docProps/app.xml", "docProps/app.xml");
                //archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/calcChain.xml", "xl/calcChain.xml");
            }

        }
    }
}
