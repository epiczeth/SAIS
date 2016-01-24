using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
namespace Sales_and_Inventory_System__Gadgets_Shop_
{
    public partial class frmLogin : Form
    {
        private const int CP_NOCLOSE_BUTTON = 0x200;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }
         String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";

        public frmLogin()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtUserName.Text == "")
            {
                MessageBox.Show("กรุณาใส่ชื่อผู้ใช้", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUserName.Focus();
                return;
            }
            if (txtPassword.Text == "")
            {
                MessageBox.Show("กรุณาใส่รหัสผ่าน", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Focus();
                return;
            }
            try
            {
                OleDbConnection myConnection = default(OleDbConnection);
                myConnection = new OleDbConnection(cs);

                OleDbCommand myCommand = default(OleDbCommand);

                myCommand = new OleDbCommand("SELECT Username,User_password FROM Users WHERE Username = @username AND User_password = @UserPassword", myConnection);
                OleDbParameter uName = new OleDbParameter("@username", OleDbType.VarChar);
                OleDbParameter uPassword = new OleDbParameter("@UserPassword", OleDbType.VarChar);
                uName.Value = txtUserName.Text;
                uPassword.Value = txtPassword.Text;
                myCommand.Parameters.Add(uName);
                myCommand.Parameters.Add(uPassword);

                myCommand.Connection.Open();

                OleDbDataReader myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection);

                if (myReader.Read() == true)
                {
                        int i;
                        ProgressBar1.Visible = true;
                        ProgressBar1.Maximum = 5000;
                        ProgressBar1.Minimum = 0;
                        ProgressBar1.Value = 4;
                        ProgressBar1.Step = 1;

                        for (i = 0; i <= 5000; i++)
                        {
                            ProgressBar1.PerformStep();
                        }
                        this.Hide();
                        frmMainMenu frm = new frmMainMenu();
                        frm.Show();
                        frm.lblUser.Text = txtUserName.Text;
                    

                    }
                

                else
                {
                    MessageBox.Show("เข้าสู่ระบบล้มเหลว...ลองใหม่อีกครั้ง !", "เข้าสู่ระบบล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    txtUserName.Clear();
                    txtPassword.Clear();
                    txtUserName.Focus();

                }
                if (myConnection.State == ConnectionState.Open)
                {
                    myConnection.Dispose();
                }

              

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
      
        private void Form1_Load(object sender, EventArgs e)
        {
            ProgressBar1.Visible = false;
            txtUserName.Focus();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
            Application.ExitThread();
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmChangePassword frm = new frmChangePassword();
            frm.Show();
            frm.txtUserName.Text = "";
            frm.txtNewPassword.Text = "";
            frm.txtOldPassword.Text = "";
            frm.txtConfirmPassword.Text = "";
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmRecoveryPassword frm = new frmRecoveryPassword();
            frm.txtEmail.Focus();
            frm.Show();
        }
    }
}
