﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Sales_and_Inventory_System__Gadgets_Shop_
{
    public partial class frmMainMenu : Form
    {
        private OleDbDataReader rdr = null;
        private OleDbConnection con = null;
        private OleDbCommand cmd = null;
        private const int CP_NOCLOSE_BUTTON = 0x200;
        protected override CreateParams CreateParams
        {
            get
            {
                var myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }
        private String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";
        public frmMainMenu()
        {
            InitializeComponent();
        }

        private void customerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmCustomers();
            frm.Show();
        }

        private void registrationToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var frm = new frmUserRegistration();
            frm.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmAbout();
            frm.Show();
        }

        private void registrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmUserRegistration();
            frm.Show();
        }

        private void profileEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmCustomers();
            frm.Show();
        }

        private void productToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmProduct();
            frm.Show();
        }

        private void notepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Notepad.exe");
        }

        private void calculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Calc.exe");
        }

        private void wordpadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Wordpad.exe");
        }

        private void taskManagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("TaskMgr.exe");
        }

        private void mSWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Winword.exe");
        }

        private void categoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmCategory();
            frm.Show();
        }

        private void companyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmCompany();
            frm.Show();
        }

        private void customersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmCustomersRecord();
            frm.Show();
        }

        private void logOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hide();
            var frm = new frmLogin();
            frm.Show();
            frm.txtUserName.Text = string.Empty;
            frm.txtPassword.Text = string.Empty;
            frm.ProgressBar1.Visible = false;
            frm.txtUserName.Focus();
        }

        private void frmMainMenu_Load(object sender, EventArgs e)
        {
            ToolStripStatusLabel4.Text = System.DateTime.Now.ToString();
            GetData();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ToolStripStatusLabel4.Text = System.DateTime.Now.ToString();
        }

        private void productsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var frm = new frmProduct();
            frm.Show();
        }

        private void productsToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            var frm = new frmProductsRecord();
            frm.Show();
        }

        private void configurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmConfig();
            frm.Show();
        }

        private void stockToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Hide();
            var frm = new frmStock();
            frm.label8.Text = lblUser.Text;
            frm.Show();
        }

        private void stockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hide();
            var frm = new frmStock();
            frm.label8.Text = lblUser.Text;
            frm.Show();
        }

        public void GetData()
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                var sql = "SELECT StockID,ProductName,Features,Price,sum(Quantity),sum(Totalprice) from Stock,Config where Stock.ConfigID=Config.ConfigID group by StockID,productname,Price,Features having sum(quantity > 0) order by ProductName";
                cmd = new OleDbCommand(sql, con);
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dataGridView1.Rows.Clear();
                while (rdr.Read() == true)
                {
                    dataGridView1.Rows.Add(rdr[0], rdr[1], rdr[2], rdr[3], rdr[4], rdr[5]);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void stockToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            var frm = new frmStockRecord();
            frm.Show();
        }

        private void invoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hide();
            var frm = new frmInvoice();
            frm.label6.Text = lblUser.Text;
            frm.Show();
        }

        private void salesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hide();
            var frm = new frmInvoice();
            frm.label6.Text = lblUser.Text;
            frm.Show();
        }

        private void salesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var frm = new frmSalesRecord();
            frm.Show();
        }

        private void loginDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new frmLoginDetails();
            frm.Show();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                var sql = "SELECT StockID,ProductName,Features,Price,sum(Quantity),sum(Totalprice) from Stock,Config where Stock.ConfigID=Config.ConfigID and ProductName like '" + txtProductName.Text + "%' group by StockID,productname,Price,Features having sum(quantity > 0) order by ProductName";
                cmd = new OleDbCommand(sql, con);
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dataGridView1.Rows.Clear();
                while (rdr.Read() == true)
                {
                    dataGridView1.Rows.Add(rdr[0], rdr[1], rdr[2], rdr[3], rdr[4], rdr[5]);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void frmMainMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }
    }
}
