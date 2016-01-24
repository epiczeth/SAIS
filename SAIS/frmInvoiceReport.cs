using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Sales_and_Inventory_System__Gadgets_Shop_
{
    public partial class frmInvoiceReport : Form
    {
        public frmInvoiceReport()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
