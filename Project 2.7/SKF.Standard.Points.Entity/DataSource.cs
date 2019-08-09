using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKF.Standard.Points.Entity
{
    public class DataSource
    {
        public static List<ComboBoxItem> Codes = new List<ComboBoxItem>()
        {
            new ComboBoxItem(){Id=0, Text= "--SELECT--"},
            new ComboBoxItem(){Id = 1, Text = "MA"},
            new ComboBoxItem(){Id = 2, Text = "MI"},
            new ComboBoxItem(){Id = 3, Text = "ME"},
            new ComboBoxItem(){Id = 4, Text = "OS"},
            new ComboBoxItem(){Id = 5, Text = "TO"},
            new ComboBoxItem(){Id = 6, Text = "DV"},
            new ComboBoxItem(){Id = 7, Text = "OI"},            
        };
    }
}
