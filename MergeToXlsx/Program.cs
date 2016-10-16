using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

namespace MergeToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("[Folder name] [Output file]");
                Console.WriteLine("[zip-file name]");
                return;
            }

            if (args.Length == 1)
            {
                using (ZipArchive archive = ZipFile.Open(args[0], ZipArchiveMode.Read))
                {
                    foreach (var entry in archive.Entries)
                    {
                        Console.WriteLine(entry.FullName);
                    }
                }

                return;
            }

            var dir_info = new DirectoryInfo(args[0]);

            if (!dir_info.Exists)
            {
                Console.WriteLine("Folder missing.");
                return;
            }

            string output_file_name = args[1];

            if (File.Exists(output_file_name))
                File.Delete(output_file_name);

            using (ZipArchive archive = ZipFile.Open(output_file_name, ZipArchiveMode.Create))
            {
                //ZipImportantFilesFirst("", dir_info, archive);
                //ZipFilesAndFolders("", dir_info, archive);
                ZipSpecificOrder("", dir_info, archive);
            }

            // var entry = os.PutNextEntry("[Content_Types].xml");
            // _rels.WriteZip(os, "_rels\\.rels");
        }

        static private void PrintFilesAndFolders(string dir_name, DirectoryInfo dir_info)
        {
            foreach(var file in dir_info.EnumerateFiles())
            {
                Console.WriteLine(dir_name + "\\" + file.Name);
            }

            foreach (var dir in dir_info.EnumerateDirectories())
            {
                Console.WriteLine(dir_name + "\\" + dir.Name);
                PrintFilesAndFolders(dir_name + "\\" + dir.Name, dir);
            }
        }

        static private void ZipFilesAndFolders(string dir_name, DirectoryInfo dir_info, ZipArchive archive)
        {
            foreach (var file in dir_info.EnumerateFiles())
            {
                if (dir_name + file.Name == "[Content_Types].xml")
                    continue;

                if (dir_name + file.Name == "_rels\\.rels")
                    continue;

                Console.WriteLine(dir_name + file.Name);

                archive.CreateEntryFromFile(file.FullName, dir_name + file.Name);
            }

            foreach (var dir in dir_info.EnumerateDirectories())
            {
                //Console.WriteLine(dir_name + "\\" + dir.Name);
                ZipFilesAndFolders(dir_name + dir.Name + "\\", dir, archive);
            }
        }

        static private void ZipImportantFilesFirst(string dir_name, DirectoryInfo dir_info, ZipArchive archive)
        {
            string content_types_xml = "[Content_Types].xml";
            Console.WriteLine(dir_name + content_types_xml);
            archive.CreateEntryFromFile(dir_info.Name + "\\" + dir_name + content_types_xml, dir_name + content_types_xml);

            string rels = "_rels\\.rels";
            Console.WriteLine(dir_name + rels);
            archive.CreateEntryFromFile(dir_info.Name + "\\" + dir_name + rels, dir_name + rels);
        }

        static private void ZipSpecificOrder(string dir_name, DirectoryInfo dir_info, ZipArchive archive)
        {
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "[Content_Types].xml", "[Content_Types].xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "_rels/.rels", "_rels/.rels");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/workbook.xml", "xl/workbook.xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/theme/theme1.xml", "xl/theme/theme1.xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/worksheets/_rels/sheet1.xml.rels", "xl/worksheets/_rels/sheet1.xml.rels");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/styles.xml", "xl/styles.xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/sharedStrings.xml", "xl/sharedStrings.xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "docProps/core.xml", "docProps/core.xml");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/printerSettings/printerSettings1.bin", "xl/printerSettings/printerSettings1.bin");
            archive.CreateEntryFromFile(dir_info.Name + "\\" + "docProps/app.xml", "docProps/app.xml");
            //archive.CreateEntryFromFile(dir_info.Name + "\\" + "xl/calcChain.xml", "xl/calcChain.xml");
        }

        /*

        */
    }
}
/*
            using(ZipArchive archive = ZipFile.Open(file_name, ZipArchiveMode.Create))
            {
                XDocument new_document = new XDocument(lxfml.SourceElement);

                ZipArchiveEntry lxfml_entry = archive.CreateEntry("IMAGE100.LXFML");
            
                using( Stream lxfml_stream = lxfml_entry.Open())
                {
                    new_document.Save(lxfml_stream); 
                }


                ZipArchiveEntry picture_entry = archive.CreateEntry("IMAGE100.PNG");

                using (Stream picture_stream = picture_entry.Open())
                {
                    picture.CopyTo(picture_stream);
                }
            }
*/
