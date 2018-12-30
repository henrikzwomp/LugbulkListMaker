using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using Newtonsoft.Json;

namespace ListMakerTwo
{
    class InputParametersWorker
    {
        public static bool ValidateParameters(InputParameters parameters, out string validation_message)
        {
            validation_message = "Parameters OK!";

            if (string.IsNullOrEmpty(parameters.SourceFileName))
            {
                validation_message = "SourceFileName not set!";
                return false;
            }

            if (!File.Exists(parameters.SourceFileName))
            {
                validation_message = "Source file not found!";
                return false;
            }

            XLWorkbook workbook = null;

            try
            {
                workbook = new XLWorkbook(parameters.SourceFileName);
            }
            catch (Exception ex)
            {
                validation_message = "Failed to open file set in SourceFileName: "
                    + ex.Message;
                return false;
            }

            if (workbook == null)
            {
                validation_message = "Failed to open file defined in SourceFileName.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.WorksheetName))
            {
                validation_message = "WorksheetName not set!";
                return false;
            }

            if (!workbook.Worksheets.Any(x => x.Name == parameters.WorksheetName))
            {
                validation_message = "Failed to find sheet defined in WorksheetName.";
                return false;
            }

            var sheet = workbook.Worksheets.First(x =>
                x.Name == parameters.WorksheetName);

            if (sheet == null)
            {
                validation_message = "Failed to open sheet defined in WorksheetName.";
                return false;
            }


            if (string.IsNullOrEmpty(parameters.BuyersSpan))
            {
                validation_message = "BuyersSpan not set!";
                return false;
            }
            if (!ValidateSpanParameter(sheet, parameters.BuyersSpan))
            {
                validation_message = "Failed to retrieve span defined in BuyersSpan.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.ElementIdSpan))
            {
                validation_message = "ElementIdSpan not set!";
                return false;
            }
            if (!ValidateSpanParameter(sheet, parameters.ElementIdSpan))
            {
                validation_message = "Failed to retrieve span defined in ElementIdSpan.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.BrickLinkDescriptionSpan))
            {
                validation_message = "BrickLinkDescriptionSpan not set!";
                return false;
            }
            if (!ValidateSpanParameter(sheet, parameters.BrickLinkDescriptionSpan))
            {
                validation_message = "Failed to retrieve span defined in BrickLinkDescriptionSpan.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.BrickLinkIdSpan))
            {
                validation_message = "BrickLinkIdSpan not set!";
                return false;
            }
            if (!ValidateSpanParameter(sheet, parameters.BrickLinkIdSpan))
            {
                validation_message = "Failed to retrieve span defined in BrickLinkIdSpan.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.BrickLinkColorSpan))
            {
                validation_message = "BrickLinkColorSpan not set!";
                return false;
            }
            if (!ValidateSpanParameter(sheet, parameters.BrickLinkColorSpan))
            {
                validation_message = "Failed to retrieve span defined in BrickLinkColorSpan.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.TlgColorSpan))
            {
                validation_message = "TlgColorSpan not set!";
                return false;
            }
            if (!ValidateSpanParameter(sheet, parameters.TlgColorSpan))
            {
                validation_message = "Failed to retrieve span defined in TlgColorSpan.";
                return false;
            }

            if (string.IsNullOrEmpty(parameters.OutputFolder))
            {
                validation_message = "SourceFileName not set!";
                return false;
            }

            if (!Directory.Exists(parameters.OutputFolder))
            {
                validation_message = "Output folder not found!";
                return false;
            }

            return true;
        }

        private static bool ValidateSpanParameter(IXLWorksheet sheet, string span)
        {
            IXLRange range = null;

            try
            {
                range = sheet.Range(span);
            }
            catch
            {

            }

            return !(range == null);
        }

        public static void CreateSettingsFile(string[] args)
        {
            var new_parameters = new InputParameters
            {
                SourceFileName = "Name of Excel file to read from",
                WorksheetName = "Name of the sheet in the Excel file to read from",
                BuyersSpan = "Range of cells to read Buyers from. Example A6:A36",
                ElementIdSpan = "Range of cells to read Element ID:s from. Example B1:Z1",
                BrickLinkDescriptionSpan = "Range of cells to read BrickLink descriptions from. Example B2:Z2",
                BrickLinkIdSpan = "Range of cells to read BrickLink ID:s from. Example B3:Z3",
                BrickLinkColorSpan = "Range of cells to read BrickLink color from. Example B4:Z4",
                TlgColorSpan = "Range of cells to read Buyers from. Example B5:Z5",
                OutputFolder = "Path to folder where all Excel files should be placed. Remember backslashes must be escaped with an extra backslash."
            };

            string json = JsonConvert.SerializeObject(new_parameters, Formatting.Indented);

            var file_name = "base_settings.txt";
            if (args.Length > 1)
                file_name = args[1];

            File.WriteAllText(file_name, json, Encoding.UTF8);
        }

        public static InputParameters ReadSettingsFile(string[] args)
        {
            InputParameters parameters;
            var json_string = File.ReadAllText(args[0]);
            parameters = JsonConvert.DeserializeObject<InputParameters>(json_string);
            return parameters;
        }

        public static InputParameters GetInputParametersFromArguments(string[] args)
        {
            return new InputParameters()
            {
                SourceFileName = args[0],
                WorksheetName = args[1],
                ElementIdSpan = args[2],
                BuyersSpan = args[3],
                BrickLinkDescriptionSpan = args[4],
                BrickLinkIdSpan = args[5],
                BrickLinkColorSpan = args[6],
                TlgColorSpan = args[7],
                OutputFolder = args[8],
            };
        }
    }
}
