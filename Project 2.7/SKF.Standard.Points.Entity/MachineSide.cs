using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace SKF.Standard.Points.Entity
{
    public static class MachineSide
    {
        public static List<ComboBoxItem> Codes = GetMachineSideCodes();
        public static List<string> EstimationCodes = GetMachineSideEstimationCodes();

        private static List<string> GetMachineSideEstimationCodes()
        {
            return new List<string>()
            {
                "NDE",
                "DE"
            };
        }

        private static List<ComboBoxItem> GetMachineSideCodes()
        {
            var list = new List<ComboBoxItem>()
            {
                new ComboBoxItem(){Id = 0, Text = "--SELECT--"},
                new ComboBoxItem(){Id = 1, Text = "NDE"},
                new ComboBoxItem(){Id = 2, Text = "DE"},
                new ComboBoxItem(){Id = 3, Text = "OB"},
                
                new ComboBoxItem(){Id = 4, Text = "IB"},
                new ComboBoxItem(){Id = 5, Text = "NDS"},
                new ComboBoxItem(){Id = 6, Text = "DS"},
                new ComboBoxItem(){Id = 7, Text = "FS"},
                
                new ComboBoxItem(){Id = 8, Text = "BS"},
                new ComboBoxItem(){Id = 9, Text = "TS"},
                new ComboBoxItem(){Id = 10, Text = "DS"},
                new ComboBoxItem(){Id = 11, Text = "GS"},
                
                new ComboBoxItem(){Id = 12, Text = "RS"},
                new ComboBoxItem(){Id = 13, Text = "PS"},
                new ComboBoxItem(){Id = 14, Text = "SBS"},
                new ComboBoxItem(){Id = 15, Text = "FWD"},
                
                new ComboBoxItem(){Id = 16, Text = "AFT"},
                new ComboBoxItem(){Id = 17, Text = "LH"},
                new ComboBoxItem(){Id = 18, Text = "RH"},
                new ComboBoxItem(){Id = 19, Text = "OU"},
 
                new ComboBoxItem(){Id = 20, Text = "IN"},
                new ComboBoxItem(){Id = 21, Text = "IPS"},
                new ComboBoxItem(){Id = 22, Text = "OPS"},
                new ComboBoxItem(){Id = 23, Text = "AVM"},
               
                new ComboBoxItem(){Id = 24, Text = "AFV"}
            };
            for (int i = 0; i <= 50; i++)
            {
                var c = $"CYL{i:D2}";
                list.Add(new ComboBoxItem(){Text = c,Id = i+25});
            }

            return list;

        }
    }
}
