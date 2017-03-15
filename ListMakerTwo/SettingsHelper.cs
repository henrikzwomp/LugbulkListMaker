using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public class SettingsHelper
    {
        public static void ReadSpan(string val, out int first_part, out int second_part)
        {
            var string_one = "";
            var string_two = "";

            ReadSpan(val, out string_one, out string_two);

            first_part = int.Parse(string_one);
            second_part = int.Parse(string_two);
        }

        public static void ReadSpan(string val, out string first_part, out string second_part)
        {
            first_part = val.Substring(0, val.IndexOf(":"));
            second_part = val.Substring(val.IndexOf(":") + 1);
        }
    }
}
