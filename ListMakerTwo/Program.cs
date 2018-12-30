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
            if (args.Length == 1 && File.Exists(args[0]))
            {
                parameters = InputParametersWorker.ReadSettingsFile(args);
            }
            if (args.Length == 9)
            {
                parameters = InputParametersWorker.
                    GetInputParametersFromArguments(args);
            }

            if (parameters == null)
            {
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
                return;
            }

            string validation_message = "";
            if(!InputParametersWorker.ValidateParameters
                (parameters, out validation_message))
            {
                Console.WriteLine(validation_message);
                return;
            }

            ExcelCreator.CreateAllExcelFiles(parameters);
        }
    }
}
