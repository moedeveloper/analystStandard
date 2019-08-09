using System.Collections.Generic;
using System.ComponentModel;

namespace SKF.Standard.Points.Entity
{
    public static class MeasurementType
    {
        public static List<string> EstimationCodes = new List<string>()
        {
            "A",
            "V",
            "E1",
            "E2",
            "E3",
            "E4"
        };

        public static List<ComboBoxItem> Codes = new List<ComboBoxItem>()
        {
            new ComboBoxItem(){Id = 0, Text = "--SELECT--"},
            new ComboBoxItem(){Id = 1, Text = "A"},
            new ComboBoxItem(){Id = 2, Text = "V"},
            new ComboBoxItem(){Id = 3, Text = "E1"},
            new ComboBoxItem(){Id = 4, Text = "E2"},
            new ComboBoxItem(){Id = 5, Text = "E3"}, // BY DEFAULT
            new ComboBoxItem(){Id = 6, Text = "E4"},
            new ComboBoxItem(){Id = 7, Text = "D"},
            new ComboBoxItem(){Id = 8, Text = "T"},
            new ComboBoxItem(){Id = 9, Text = "G"},
            new ComboBoxItem(){Id = 10, Text = "P"},
            new ComboBoxItem(){Id = 11, Text = "B"},
            new ComboBoxItem(){Id = 12, Text = "S"},

            new ComboBoxItem(){Id = 13, Text = "EC"},
            new ComboBoxItem(){Id = 14, Text = "AE"},
            new ComboBoxItem(){Id = 15, Text = "HD"}, // BY DEFAULT
            new ComboBoxItem(){Id = 16, Text = "MD"},

            new ComboBoxItem(){Id = 17, Text = "PD"},
            new ComboBoxItem(){Id = 18, Text = "PS"},
            new ComboBoxItem(){Id = 19, Text = "SG"}, // BY DEFAULT
            new ComboBoxItem(){Id = 20, Text = "TT"},

            new ComboBoxItem(){Id = 21, Text = "TO"},
            new ComboBoxItem(){Id = 22, Text = "VT"},
            new ComboBoxItem(){Id = 23, Text = "CT"}, // BY DEFAULT
            new ComboBoxItem(){Id = 24, Text = "MC"},
            new ComboBoxItem(){Id = 25, Text = "LT"},


        };

    }
}
