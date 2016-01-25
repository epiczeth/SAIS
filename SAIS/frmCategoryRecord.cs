﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace SAIS
{
    public partial class frmCategoryRecord : Form
    {
        private OleDbDataReader rdr = null;
        private OleDbConnection con = null;
        private OleDbCommand cmd = null;
        private String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";

        public frmCategoryRecord()
        {
            InitializeComponent();
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
            var dr = dataGridView1.SelectedRows[0];
            Hide();
            var frm = new frmCategory();


            frm.Show();
            frm.txtCategoryName.Text = dr.Cells[0].Value.ToString();
            frm.textBox1.Text = dr.Cells[0].Value.ToString();
            frm.btnDelete.Enabled = true;
            frm.btnUpdate.Enabled = true;
            frm.txtCategoryName.Focus();
            frm.btnSave.Enabled = false;
        }
        public void GetData()
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                var sql = "SELECT * from Category order by CategoryName";
                cmd = new OleDbCommand(sql, con);
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dataGridView1.Rows.Clear();
                while (rdr.Read() == true)
                {
                    dataGridView1.Rows.Add(rdr[0]);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void frmCategoryRecord_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            var frm = new frmCategory();
            frm.Show();
        }

        private void frmCategoryRecord_Load(object sender, EventArgs e)
        {
            GetData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var dr = dataGridView1.SelectedRows[0];
                Hide();
                var frm = new frmCategory();


                frm.Show();
                frm.txtCategoryName.Text = dr.Cells[0].Value.ToString();
                frm.textBox1.Text = dr.Cells[0].Value.ToString();
                frm.btnDelete.Enabled = true;
                frm.btnUpdate.Enabled = true;
                frm.txtCategoryName.Focus();
                frm.btnSave.Enabled = false;
            }
            catch
            {
            }
        }
    }
}
