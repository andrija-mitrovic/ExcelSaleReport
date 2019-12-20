using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelSaleReport
{
    public static class Validation
    {
        public static void NumberInterval(object sender, CancelEventArgs e)
        {
            int num;
            TextBox tb = (TextBox)sender;
            if (Int32.TryParse(tb.Text, out num) && num > 0 && num < 13)
                e.Cancel = false;
            else
            {
                e.Cancel = true;
                MessageBox.Show("Number must be between 1 and 12", "Attention!");
            }
        }
    }
}
