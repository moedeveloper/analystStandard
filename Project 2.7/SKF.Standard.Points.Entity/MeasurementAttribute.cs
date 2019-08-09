using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKF.Standard.Points.Entity
{
    public static class MeasurementAttribute
    {
        public static List<ComboBoxItem> Codes = new List<ComboBoxItem>()
        {
            new ComboBoxItem(){Id = 0, Text = "--SELECT--"},
            new ComboBoxItem(){Id = 0, Text = "LF"},
            new ComboBoxItem(){Id = 0, Text = "MF"},
            new ComboBoxItem(){Id = 0, Text = "HF"},
            new ComboBoxItem(){Id = 0, Text = "HC"},
            new ComboBoxItem(){Id = 0, Text = "MC"},
            new ComboBoxItem(){Id = 0, Text = "LC"},
            new ComboBoxItem(){Id = 0, Text = "MR"},
            new ComboBoxItem(){Id = 0, Text = "LR"},
            new ComboBoxItem(){Id = 0, Text = "OR"},
            new ComboBoxItem(){Id = 0, Text = "SC"},
            new ComboBoxItem(){Id = 0, Text = "TW"}

        };
    }
}
