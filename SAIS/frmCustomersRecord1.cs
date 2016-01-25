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
    public partial class frmCustomersRecord1 : Form
    {
        private OleDbConnection con = null;
        private OleDbCommand cmd = null;
        private String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";
        public frmCustomersRecord1()
        {
            InitializeComponent();
        }
        public void GetData()
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                cmd = new OleDbCommand( "SELECT (CustomerID)as [Customer ID],(Customername) as [Customer Name],(address) as [Address],(landmark) as [Landmark],(city) as [City],(state) as [State],(zipcode) as [Zip/Post Code],(Phone) as [Phone],(email) as [Email],(mobileno) as [Mobile No],(faxno) as [Fax No],(notes) as [Notes] from Customer order by CustomerName", con);
                var myDA = new OleDbDataAdapter(cmd);
                var myDataSet = new DataSet();
                myDA.Fill(myDataSet, "Customer");
                dataGridView1.DataSource = myDataSet.Tables["Customer"].DefaultView;
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void frmCustomersRecord_Load(object sender, EventArgs e)
        {
            GetData();
        }



        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var strRowNumber = (e.RowIndex + 1).ToString();
            var size = e.Graphics.MeasureString(strRowNumber, Font);
            if (dataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)))
            {
                dataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20));
            }
            var b = SystemBrushes.ControlText;
            e.Graphics.DrawString(strRowNumber, Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2));
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var dr = dataGridView1.SelectedRows[0];
                Hide();
                var frm = new frmInvoice();
                frm.Show();
                frm.txtCustomerID.Text = dr.Cells[0].Value.ToString();
                frm.txtCustomerName.Text = dr.Cells[1].Value.ToString();
                frm.label6.Text = label1.Text;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtCustomers_TextChanged(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                cmd = new OleDbCommand("SELECT (CustomerID)as [Customer ID],(Customername) as [Customer Name],(address) as [Address],(landmark) as [Landmark],(city) as [City],(state) as [State],(zipcode) as [Zip/Post Code],(Phone) as [Phone],(email) as [Email],(mobileno) as [Mobile No],(faxno) as [Fax No],(notes) as [Notes] from Customer where CustomerName like '" + txtCustomers.Text + "%' order by CustomerName", con);
                var myDA = new OleDbDataAdapter(cmd);
                var myDataSet = new DataSet();
                myDA.Fill(myDataSet, "Customer");
                dataGridView1.DataSource = myDataSet.Tables["Customer"].DefaultView;
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource == null)
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

                rowsTotal = dataGridView1.RowCount - 1;
                colsTotal = dataGridView1.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = dataGridView1.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = dataGridView1.Rows[I].Cells[j].Value;
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

        private void frmCustomersRecord1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            var frm = new frmInvoice();
            frm.label6.Text = label1.Text;
            frm.Show();
        }
    }
}
