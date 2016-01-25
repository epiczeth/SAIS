﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace SAIS
{
    public partial class frmLogin : Form
    {
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

        public frmLogin()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtUserName.Text == string.Empty)
            {
                MessageBox.Show("กรุณาใส่ชื่อผู้ใช้", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUserName.Focus();
                return;
            }
            if (txtPassword.Text == string.Empty)
            {
                MessageBox.Show("กรุณาใส่รหัสผ่าน", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Focus();
                return;
            }
            try
            {
                var myConnection = default(OleDbConnection);
                myConnection = new OleDbConnection(cs);

                var myCommand = default(OleDbCommand);

                myCommand = new OleDbCommand("SELECT Username,User_password FROM Users WHERE Username = @username AND User_password = @UserPassword", myConnection);
                var uName = new OleDbParameter("@username", OleDbType.VarChar);
                var uPassword = new OleDbParameter("@UserPassword", OleDbType.VarChar);
                uName.Value = txtUserName.Text;
                uPassword.Value = txtPassword.Text;
                myCommand.Parameters.Add(uName);
                myCommand.Parameters.Add(uPassword);

                myCommand.Connection.Open();

                var myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection);

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
                    Hide();
                    var frm = new frmMainMenu();
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
            Dispose();
            Application.ExitThread();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            var frm = new frmChangePassword();
            frm.Show();
            frm.txtUserName.Text = string.Empty;
            frm.txtNewPassword.Text = string.Empty;
            frm.txtOldPassword.Text = string.Empty;
            frm.txtConfirmPassword.Text = string.Empty;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            var frm = new frmRecoveryPassword();
            frm.txtEmail.Focus();
            frm.Show();
        }
    }
}
