using System.Collections.Generic;
using System.ComponentModel;

namespace SKF.Standard.Points.Entity
{
    public static class AngularOrientation
    {
        public static List<ComboBoxItem> Codes = GetCodeList();
        public static List<string> EstimationCodes = GetEstimationCodeList();

        private static List<string> GetEstimationCodeList()
        {
            return new List<string>()
            {
                "H",
                "V",
                "A",
                "X",
                "Z",
                "XY",
                "XYZ"
            };
        }

        private static List<ComboBoxItem> GetCodeList()
        {
            List<ComboBoxItem> list = new List<ComboBoxItem>()
            {
                new ComboBoxItem(){Id = 0, Text = "--SELECT--"},
                new ComboBoxItem(){Id = 1, Text = "H"},
                new ComboBoxItem(){Id = 2, Text = "V"},
                new ComboBoxItem(){Id = 3, Text = "A"},
                new ComboBoxItem(){Id = 4, Text = "R"}
            };

            string radial = "";
            int j = 15;
            var sum = 0;
            for (int i = 0; i < 24; i++)
            {
                if (i == 0)
                {
                    radial = $"R{i:D3}";
                }
                else
                {
                    sum += j;
                    radial = $"R{sum:D3}";
                }
                list.Add(new ComboBoxItem(){Text = radial, Id = i+5});
            }

            return list;
        }
        
    }
}
