using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace MarkingSheet
{
    internal class ExclusionList
    {

        public static List<String> excludedIsins()
        {
            return new List<string>()
            {
                "XS1593171967_OLD_230706",
                "Project Gilbert",
                "PROJECT EARTH LOAN",
                "Kaisag loan 01/06/22",
                "Harper \"Lucas Drilling\"",
                "Brompton",
                "ACCIL Claim #4 -2",
                "ACCIL Claim #3 - 2",
                "ACCIL Claim #2 -2 ",
                "ACCIL Claim #1 - 2",
                "USQ82780AG49",
            };
        }

    }
}
