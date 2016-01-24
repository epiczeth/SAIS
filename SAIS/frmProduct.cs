using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Sales_and_Inventory_System__Gadgets_Shop_
{
    public partial class frmProduct : Form
    {
        private OleDbDataReader rdr = null;
        private OleDbConnection con = null;
        private OleDbCommand cmd = null;
        private String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";

        public frmProduct()
        {
            InitializeComponent();
        }

        private void frmProduct_Load(object sender, EventArgs e)
        {
            FillCombo();
            Autocomplete();
        }
        public void FillCombo()
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                var ct = "select RTRIM(CategoryName) from Category order by CategoryName";
                cmd = new OleDbCommand(ct);
                cmd.Connection = con;
                rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    cmbCategory.Items.Add(rdr[0]);
                }
                con.Close();
                con = new OleDbConnection(cs);
                con.Open();
                var ct1 = "select RTRIM(CompanyName) from Company order by CompanyName";
                cmd = new OleDbCommand(ct1);
                cmd.Connection = con;
                rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    cmbCompany.Items.Add(rdr[0]);
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Reset()
        {
            txtProductName.Text = string.Empty;
            cmbCompany.Text = string.Empty;
            cmbCategory.Text = string.Empty;
            btnDelete.Enabled = false;
            btnUpdate.Enabled = false;
            btnSave.Enabled = true;
            txtProductName.Focus();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Reset();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtProductName.Text == string.Empty)
            {
                MessageBox.Show("Please enter product name", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtProductName.Focus();
                return;
            }
            if (cmbCategory.Text == string.Empty)
            {
                MessageBox.Show("Please select category", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbCategory.Focus();
                return;
            }
            if (cmbCompany.Text == string.Empty)
            {
                MessageBox.Show("Please select company", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbCompany.Focus();
                return;
            }

            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                var ct = "select ProductName from Product where ProductName='" + txtProductName.Text + "'";

                cmd = new OleDbCommand(ct);
                cmd.Connection = con;
                rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    MessageBox.Show("Product Name Already Exists", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtProductName.Text = string.Empty;
                    txtProductName.Focus();


                    if ((rdr != null))
                    {
                        rdr.Close();
                    }
                    return;
                }

                con = new OleDbConnection(cs);
                con.Open();

                var cb = "insert into Product(ProductName,Category,Company) VALUES ('" + txtProductName.Text + "','" + cmbCategory.Text + "','" + cmbCompany.Text + "')";
                cmd = new OleDbCommand(cb);
                cmd.Connection = con;
                cmd.ExecuteReader();
                con.Close();
                MessageBox.Show("บันทึกข้อมูลสำเร็จ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Autocomplete();
                txtProductName.Text = string.Empty;
                txtProductName.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("คุณต้องการลบข้อมูลนี้จริงหรือไม่?", "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                delete_records();
            }
        }
        private void delete_records()
        {
            try
            {
                var RowsAffected = 0;
                con = new OleDbConnection(cs);
                con.Open();
                var cq = "delete from product where productName='" + txtProductName.Text + "'";
                cmd = new OleDbCommand(cq);
                cmd.Connection = con;
                RowsAffected = cmd.ExecuteNonQuery();
                if (RowsAffected > 0)
                {
                    MessageBox.Show("ลบข้อมูลสำเร็จ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Reset();
                    Autocomplete();
                }
                else
                {
                    MessageBox.Show("ข้อมูลดังกล่าวไม่มีอยู่จริง", "ขออภัย", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Reset();
                    Autocomplete();
                }
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Autocomplete()
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();
                var cmd = new OleDbCommand("SELECT distinct ProductName FROM product", con);
                var ds = new DataSet();
                var da = new OleDbDataAdapter(cmd);
                da.Fill(ds, "Product");
                var col = new AutoCompleteStringCollection();
                var i = 0;
                for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                {
                    col.Add(ds.Tables[0].Rows[i]["productname"].ToString());
                }
                txtProductName.AutoCompleteSource = AutoCompleteSource.CustomSource;
                txtProductName.AutoCompleteCustomSource = col;
                txtProductName.AutoCompleteMode = AutoCompleteMode.Suggest;

                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection(cs);
                con.Open();

                var cb = "update Product set productName='" + txtProductName.Text + "',Category='" + cmbCategory.Text + "', Company='" + cmbCompany.Text + "' where Productname='" + textBox1.Text + "'";
                cmd = new OleDbCommand(cb);
                cmd.Connection = con;
                cmd.ExecuteReader();
                con.Close();
                MessageBox.Show("ปรับปรุงข้อมูลสำเร็จ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Autocomplete();
                btnUpdate.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
            Hide();
            var frm = new frmProductsRecord1();
            frm.Show();
            frm.GetData();
        }
    }
}
