using ExcelSaleReport.Data;
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
    public partial class Hour : Form
    {
        public Hour()
        {
            InitializeComponent();

            this.MinimizeBox = false;
            this.MaximizeBox = false;

            this.accept.Click += Accept_Click;
        }

        private void Accept_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            this.accept.Enabled = false;

            Reports reports = new Reports(new ProductRepository());
            reports.GetProductTypeRealizationByHour(
                Convert.ToDateTime(this.dateFrom.Text),
                Convert.ToDateTime(this.dateTo.Text));

            this.accept.Enabled = true;
            this.Close();
        }
    }
}
