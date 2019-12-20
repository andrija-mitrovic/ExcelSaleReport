using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelSaleReport
{
    public partial class Menu : Form
    {
        private Form _form;
        public Menu()
        {
            InitializeComponent();

            this.MinimizeBox = false;
            this.MaximizeBox = false;

            this.b_hour.Click += Click_Hour;
            this.b_day.Click += Click_Day;
            this.b_supplier.Click += Click_Supplier;
        }

        private void Click_Supplier(object sender, EventArgs e)
        {
            _form = new Supplier();
            _form.ShowDialog();
        }

        private void Click_Day(object sender, EventArgs e)
        {
            _form = new Day();
            _form.ShowDialog();
        }

        private void Click_Hour(object sender, EventArgs e)
        {
            _form = new Hour();
            _form.ShowDialog();
        }
    }
}
