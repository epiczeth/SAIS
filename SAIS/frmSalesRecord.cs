﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAIS
{
    public partial class frmSalesRecord : Form
    {
        private DataTable dTable;
        private OleDbConnection con = null;
        private OleDbDataAdapter adp;
        private DataSet ds;
        private OleDbCommand cmd = null;
        private String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";
        public frmSalesRecord()
        {
            InitializeComponent();
        }

        private void frmSalesRecord_Load(object sender, EventArgs e)
        {
            FillCombo();
            TabControl1.TabPages.Remove(TabPage1);
            TabControl1.TabPages.Remove(TabPage2);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (DataGridView1.DataSource == null)
            {
                MessageBox.Show("ไม่มีข้อมูลในการสร้างไฟล์ Excel", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var rowsTotal = 0;
            var colsTotal = 0;
            var I = 0;
            var j = 0;
            var iC = 0;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            var xlApp = new Excel.Application();

            try
            {
                var excelBook = xlApp.Workbooks.Add();
                var excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[1];
                xlApp.Visible = true;

                rowsTotal = DataGridView1.RowCount - 1;
                colsTotal = DataGridView1.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = DataGridView1.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = DataGridView1.Rows[I].Cells[j].Value;
                    }
                }
                _with1.Rows["1:1"].Font.FontStyle = "Bold";
                _with1.Rows["1:1"].Font.Size = 12;

                _with1.Cells.Columns.AutoFit();
                _with1.Cells.Select();
                _with1.Cells.EntireColumn.AutoFit();
                _with1.Cells[1, 1].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                xlApp = null;
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            if (DataGridView2.DataSource == null)
            {
                MessageBox.Show("ไม่มีข้อมูลในการสร้างไฟล์ Excel", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var rowsTotal = 0;
            var colsTotal = 0;
            var I = 0;
            var j = 0;
            var iC = 0;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            var xlApp = new Excel.Application();

            try
            {
                var excelBook = xlApp.Workbooks.Add();
                var excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[1];
                xlApp.Visible = true;

                rowsTotal = DataGridView2.RowCount - 1;
                colsTotal = DataGridView2.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = DataGridView2.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = DataGridView2.Rows[I].Cells[j].Value;
                    }
                }
                _with1.Rows["1:1"].Font.FontStyle = "Bold";
                _with1.Rows["1:1"].Font.Size = 12;

                _with1.Cells.Columns.AutoFit();
                _with1.Cells.Select();
                _with1.Cells.EntireColumn.AutoFit();
                _with1.Cells[1, 1].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                xlApp = null;
            }
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            if (DataGridView3.DataSource == null)
            {
                MessageBox.Show("ไม่มีข้อมูลในการสร้างไฟล์ Excel", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var rowsTotal = 0;
            var colsTotal = 0;
            var I = 0;
            var j = 0;
            var iC = 0;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            var xlApp = new Excel.Application();

            try
            {
                var excelBook = xlApp.Workbooks.Add();
                var excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[1];
                xlApp.Visible = true;

                rowsTotal = DataGridView3.RowCount - 1;
                colsTotal = DataGridView3.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = DataGridView3.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = DataGridView3.Rows[I].Cells[j].Value;
                    }
                }
                _with1.Rows["1:1"].Font.FontStyle = "Bold";
                _with1.Rows["1:1"].Font.Size = 12;

                _with1.Cells.Columns.AutoFit();
                _with1.Cells.Select();
                _with1.Cells.EntireColumn.AutoFit();
                _with1.Cells[1, 1].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                xlApp = null;
            }
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            DataGridView3.DataSource = null;
            cmbCustomerName.Text = string.Empty;
            GroupBox4.Visible = false;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            DataGridView1.DataSource = null;
            dtpInvoiceDateFrom.Text = DateTime.Today.ToString();
            dtpInvoiceDateTo.Text = DateTime.Today.ToString();
            GroupBox3.Visible = false;
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            DateTimePicker1.Text = DateTime.Today.ToString();
            DateTimePicker2.Text = DateTime.Today.ToString();
            DataGridView2.DataSource = null;
            GroupBox10.Visible = false;
        }
        public void FillCombo()
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                adp = new OleDbDataAdapter();
                adp.SelectCommand = new OleDbCommand("SELECT distinct CustomerName FROM Sales,Customer where Sales.CustomerID=Customer.CustomerID", con);
                ds = new DataSet("ds");
                adp.Fill(ds);
                dTable = ds.Tables[0];
                cmbCustomerName.Items.Clear();
                foreach (DataRow drow in dTable.Rows)
                {
                    cmbCustomerName.Items.Add(drow[0].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                GroupBox3.Visible = true;
                con = new OleDbConnection(cs);
                con.Open();
                cmd = new OleDbCommand("SELECT (invoiceNo) as [Invoice No],(InvoiceDate) as [Invoice Date],(Sales.CustomerID) as [Customer ID],(CustomerName) as [Customer Name],(GrandTotal) as [Grand Total],(TotalPayment) as [Total Payment],(PaymentDue) as [Payment Due] from Sales,Customer where Sales.CustomerID=Customer.CustomerID and InvoiceDate between #" + dtpInvoiceDateFrom.Text + "# And #" + dtpInvoiceDateTo.Text + "# order by InvoiceDate desc", con);
                var myDA = new OleDbDataAdapter(cmd);
                var myDataSet = new DataSet();
                myDA.Fill(myDataSet, "Sales");
                myDA.Fill(myDataSet, "Customer");
                DataGridView1.DataSource = myDataSet.Tables["Customer"].DefaultView;
                DataGridView1.DataSource = myDataSet.Tables["Sales"].DefaultView;
                Int64 sum = 0;
                Int64 sum1 = 0;
                Int64 sum2 = 0;

                foreach (DataGridViewRow r in DataGridView1.Rows)
                {
                    var i = Convert.ToInt64(r.Cells[4].Value);
                    var j = Convert.ToInt64(r.Cells[5].Value);
                    var k = Convert.ToInt64(r.Cells[6].Value);
                    sum = sum + i;
                    sum1 = sum1 + j;
                    sum2 = sum2 + k;
                }
                TextBox1.Text = sum.ToString();
                TextBox2.Text = sum1.ToString();
                TextBox3.Text = sum2.ToString();

                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbCustomerName_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                GroupBox4.Visible = true;
                con = new OleDbConnection(cs);
                con.Open();
                cmd = new OleDbCommand("SELECT (invoiceNo) as [Invoice No],(InvoiceDate) as [Invoice Date],(Sales.CustomerID) as [Customer ID],(CustomerName) as [Customer Name],(GrandTotal) as [Grand Total],(TotalPayment) as [Total Payment],(PaymentDue) as [Payment Due] from Sales,Customer where Sales.CustomerID=Customer.CustomerID and Customername='" + cmbCustomerName.Text + "' order by CustomerName,InvoiceDate", con);
                var myDA = new OleDbDataAdapter(cmd);
                var myDataSet = new DataSet();
                myDA.Fill(myDataSet, "Sales");
                myDA.Fill(myDataSet, "Customer");
                DataGridView3.DataSource = myDataSet.Tables["Customer"].DefaultView;
                DataGridView3.DataSource = myDataSet.Tables["Sales"].DefaultView;
                Int64 sum = 0;
                Int64 sum1 = 0;
                Int64 sum2 = 0;
                var header = new string[7]
                { "รหัสสั่งซื้อ", "วันที่สั่งซื้อ", "รหัสลูกค้า", "ชื่อลูกค้า", "ราคารวมภาษี", "รวมเป็นเงิน", "ส่วนลด"
                };
                for (var i = 0; i < header.Length; i++)
                {
                    DataGridView3.Columns[i].HeaderText = header[i];
                }
                foreach (DataGridViewRow r in DataGridView3.Rows)
                {
                    var i = Convert.ToInt64(r.Cells[4].Value);
                    var j = Convert.ToInt64(r.Cells[5].Value);
                    var k = Convert.ToInt64(r.Cells[6].Value);
                    sum = sum + i;
                    sum1 = sum1 + j;
                    sum2 = sum2 + k;
                }
                TextBox6.Text = sum.ToString();
                TextBox5.Text = sum1.ToString();
                TextBox4.Text = sum2.ToString();

                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            try
            {
                GroupBox10.Visible = true;
                con = new OleDbConnection(cs);
                con.Open();
                cmd = new OleDbCommand("SELECT (invoiceNo) as [Invoice No],(InvoiceDate) as [Invoice Date],(Sales.CustomerID) as [Customer ID],(CustomerName) as [Customer Name],(GrandTotal) as [Grand Total],(TotalPayment) as [Total Payment],(PaymentDue) as [Payment Due] from Sales,Customer where Sales.CustomerID=Customer.CustomerID and InvoiceDate between #" + DateTimePicker2.Text + "# And #" + DateTimePicker1.Text + "# and PaymentDue > 0 order by InvoiceDate desc", con);
                var myDA = new OleDbDataAdapter(cmd);
                var myDataSet = new DataSet();
                myDA.Fill(myDataSet, "Sales");
                myDA.Fill(myDataSet, "Customer");
                DataGridView2.DataSource = myDataSet.Tables["Customer"].DefaultView;
                DataGridView2.DataSource = myDataSet.Tables["Sales"].DefaultView;
                Int64 sum = 0;
                Int64 sum1 = 0;
                Int64 sum2 = 0;

                foreach (DataGridViewRow r in DataGridView2.Rows)
                {
                    var i = Convert.ToInt64(r.Cells[4].Value);
                    var j = Convert.ToInt64(r.Cells[5].Value);
                    var k = Convert.ToInt64(r.Cells[6].Value);
                    sum = sum + i;
                    sum1 = sum1 + j;
                    sum2 = sum2 + k;
                }
                TextBox12.Text = sum.ToString();
                TextBox11.Text = sum1.ToString();
                TextBox10.Text = sum2.ToString();

                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void DataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var strRowNumber = (e.RowIndex + 1).ToString();
            var size = e.Graphics.MeasureString(strRowNumber, Font);
            if (DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)))
            {
                DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20));
            }
            var b = SystemBrushes.ControlText;
            e.Graphics.DrawString(strRowNumber, Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2));
        }


        private void DataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var strRowNumber = (e.RowIndex + 1).ToString();
            var size = e.Graphics.MeasureString(strRowNumber, Font);
            if (DataGridView3.RowHeadersWidth < Convert.ToInt32((size.Width + 20)))
            {
                DataGridView3.RowHeadersWidth = Convert.ToInt32((size.Width + 20));
            }
            var b = SystemBrushes.ControlText;
            e.Graphics.DrawString(strRowNumber, Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2));
        }

        private void DataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var strRowNumber = (e.RowIndex + 1).ToString();
            var size = e.Graphics.MeasureString(strRowNumber, Font);
            if (DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)))
            {
                DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20));
            }
            var b = SystemBrushes.ControlText;
            e.Graphics.DrawString(strRowNumber, Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2));
        }


        private void TabControl1_Click(object sender, EventArgs e)
        {
            DataGridView1.DataSource = null;
            dtpInvoiceDateFrom.Text = DateTime.Today.ToString();
            dtpInvoiceDateTo.Text = DateTime.Today.ToString();
            GroupBox3.Visible = false;
            DataGridView3.DataSource = null;
            cmbCustomerName.Text = string.Empty;
            GroupBox4.Visible = false;
            DateTimePicker1.Text = DateTime.Today.ToString();
            DateTimePicker2.Text = DateTime.Today.ToString();
            DataGridView2.DataSource = null;
            GroupBox10.Visible = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                timer1.Enabled = true;
                var rpt = new rptSales();

                cmd = new OleDbCommand();
                var myDA = new OleDbDataAdapter();
                var myDS = new SIS_DBDataSet();

                con = new OleDbConnection(cs);
                cmd.Connection = con;
                cmd.CommandText = "SELECT Sales.InvoiceNo, Sales.InvoiceDate, Sales.CustomerID, Sales.SubTotal, Sales.VATPercentage, Sales.VATAmount, Sales.GrandTotal, Sales.TotalPayment,Sales.PaymentDue, Sales.Remarks, Customer.CustomerID AS Expr1, Customer.CustomerName, Customer.Address, Customer.Landmark, Customer.City, Customer.State, Customer.ZipCode, Customer.Phone, Customer.MobileNo, Customer.FaxNo, Customer.Email, Customer.Notes FROM (Sales INNER JOIN Customer ON Sales.CustomerID = Customer.CustomerID) where InvoiceDate between #" + dtpInvoiceDateFrom.Text + "# And #" + dtpInvoiceDateTo.Text + "# order by InvoiceDate desc";
                cmd.CommandType = CommandType.Text;
                myDA.SelectCommand = cmd;
                myDA.Fill(myDS, "Sales");
                myDA.Fill(myDS, "Customer");
                rpt.SetDataSource(myDS);
                var frm = new frmSalesReport();
                frm.crystalReportViewer1.ReportSource = rpt;
                frm.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
            timer1.Enabled = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbCustomerName.Text == string.Empty)
                {
                    MessageBox.Show("Please select Customer name", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cmbCustomerName.Focus();
                    return;
                }
                Cursor = Cursors.WaitCursor;
                timer1.Enabled = true;

                var rpt = new rptSales();

                cmd = new OleDbCommand();
                var myDA = new OleDbDataAdapter();
                var myDS = new SIS_DBDataSet();

                con = new OleDbConnection(cs);
                cmd.Connection = con;
                cmd.CommandText = "SELECT Sales.InvoiceNo, Sales.InvoiceDate, Sales.CustomerID, Sales.SubTotal, Sales.VATPercentage, Sales.VATAmount, Sales.GrandTotal, Sales.TotalPayment,Sales.PaymentDue, Sales.Remarks, Customer.CustomerID AS Expr1, Customer.CustomerName, Customer.Address, Customer.Landmark, Customer.City, Customer.State, Customer.ZipCode, Customer.Phone, Customer.MobileNo, Customer.FaxNo, Customer.Email, Customer.Notes FROM (Sales INNER JOIN Customer ON Sales.CustomerID = Customer.CustomerID) where Customername='" + cmbCustomerName.Text + "' order by CustomerName,InvoiceDate";
                cmd.CommandType = CommandType.Text;
                myDA.SelectCommand = cmd;
                myDA.Fill(myDS, "Sales");
                myDA.Fill(myDS, "Customer");
                rpt.SetDataSource(myDS);
                var frm = new frmSalesReport();
                frm.crystalReportViewer1.ReportSource = rpt;
                frm.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                timer1.Enabled = true;

                var rpt = new rptSales();

                cmd = new OleDbCommand();
                var myDA = new OleDbDataAdapter();
                var myDS = new SIS_DBDataSet();

                con = new OleDbConnection(cs);
                cmd.Connection = con;
                cmd.CommandText = "SELECT Sales.InvoiceNo, Sales.InvoiceDate, Sales.CustomerID, Sales.SubTotal, Sales.VATPercentage, Sales.VATAmount, Sales.GrandTotal, Sales.TotalPayment,Sales.PaymentDue, Sales.Remarks, Customer.CustomerID AS Expr1, Customer.CustomerName, Customer.Address, Customer.Landmark, Customer.City, Customer.State, Customer.ZipCode, Customer.Phone, Customer.MobileNo, Customer.FaxNo, Customer.Email, Customer.Notes FROM (Sales INNER JOIN Customer ON Sales.CustomerID = Customer.CustomerID) Where InvoiceDate between #" + DateTimePicker2.Text + "# And #" + DateTimePicker1.Text + "# and PaymentDue > 0 order by InvoiceDate desc";
                cmd.CommandType = CommandType.Text;
                myDA.SelectCommand = cmd;
                myDA.Fill(myDS, "Sales");
                myDA.Fill(myDS, "Customer");
                rpt.SetDataSource(myDS);
                var frm = new frmSalesReport();
                frm.crystalReportViewer1.ReportSource = rpt;
                frm.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
