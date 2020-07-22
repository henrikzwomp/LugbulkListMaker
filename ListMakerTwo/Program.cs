using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

namespace ListMakerTwo
{
    class Program
    {
        static void Main(string[] args)
        {
            InputParameters parameters = null;

            if (args.Length >= 1 && string.Equals(args[0], "--CreateSettingsFile"
                , StringComparison.OrdinalIgnoreCase))
            {
                InputParametersWorker.CreateSettingsFile(args);
                return;
            }
            else if (args.Length == 3 && string.Equals(args[0], "--CreateSqlFile"
                , StringComparison.OrdinalIgnoreCase))
            {
                parameters = InputParametersWorker.ReadSettingsFile(args[1]);
                var sheet = SheetRetriever.Get(parameters.SourceFileName,
                parameters.WorksheetName);
                var reader = new SourceReader(sheet, parameters);
                var sql_lines = SqlFileCreator.MakeFileForLugbulkDatabase(reader);
                File.WriteAllLines(args[2], sql_lines);
                return;
            }
            else if (args.Length == 4 && string.Equals(args[0], "--UseLists"
                , StringComparison.OrdinalIgnoreCase))
            {
                var orders = File.ReadAllLines(args[1]).ToList();
                var elements = File.ReadAllLines(args[2]).ToList();

                var reader = new CsvFileSourceReader(orders, elements);

                ExcelCreator.CreateAllExcelFiles(reader, args[3]);

                return;
            }
            else if (args.Length == 1 && File.Exists(args[0]))
            {
                parameters = InputParametersWorker.ReadSettingsFile(args[0]);
            }
            else if (args.Length == 9)
            {
                parameters = InputParametersWorker.
                    GetInputParametersFromArguments(args);
            }

            if (parameters != null)
            {
                string validation_message = "";
                if (!InputParametersWorker.ValidateParameters
                    (parameters, out validation_message))
                {
                    Console.WriteLine(validation_message);
                    return;
                }

                var sheet = SheetRetriever.Get(parameters.SourceFileName,
                    parameters.WorksheetName);

                var reader = new SourceReader(sheet, parameters);

                ExcelCreator.CreateAllExcelFiles(reader, parameters.OutputFolder);

                return;
            }

            Console.WriteLine("Commands: ");
            Console.WriteLine("");
            Console.WriteLine("ListMakerTwo.exe <Source File Name> <Sheet Name> " +
                "<ElementId Span> <Buyers Span> <BrickLink Description Span> " +
                "<BrickLinkId Span> <BrickLink Color Span> <TLG Color Span>" +
                "<output folder>");
            Console.WriteLine("");
            Console.WriteLine("ListMakerTwo.exe <settings file>");
            Console.WriteLine("");
            Console.WriteLine("ListMakerTwo.exe --CreateSettingsFile [Optional settings file name]");
            Console.WriteLine("(Default file name will be base_settings.txt)");
            Console.WriteLine("");
            Console.WriteLine("ListMakerTwo.exe --CreateSqlFile <settings file> <output file>");
            Console.WriteLine("");
            Console.WriteLine("ListMakerTwo.exe --UseLists <orders data file> <element data file> <output folder>");

        }
    }
}
