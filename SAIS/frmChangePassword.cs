﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Sales_and_Inventory_System__Gadgets_Shop_
{
    public partial class frmChangePassword : Form
    {
        private OleDbConnection con = null;
        private OleDbCommand cmd = null;
        private String cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\SIS_DB.accdb;";

        public frmChangePassword()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                var RowsAffected = 0;
                if ((txtUserName.Text.Trim().Length == 0))
                {
                    MessageBox.Show("กรุณาระบุชื่อผู้ใช้", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtUserName.Focus();
                    return;
                }
                if ((txtOldPassword.Text.Trim().Length == 0))
                {
                    MessageBox.Show("กรุณาระบุรหัสผ่านเก่า", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtOldPassword.Focus();
                    return;
                }
                if ((txtNewPassword.Text.Trim().Length == 0))
                {
                    MessageBox.Show("กรุณาระบุรหัสผ่านใหม่", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtNewPassword.Focus();
                    return;
                }
                if ((txtConfirmPassword.Text.Trim().Length == 0))
                {
                    MessageBox.Show("กรุณาระบุรหัสผ่านใหม่อีกครั้ง", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtConfirmPassword.Focus();
                    return;
                }
                if ((txtNewPassword.TextLength < 5))
                {
                    MessageBox.Show("รหัสผ่านใหม่ต้องมีความยาม 5 ตัวอักษรขึ้นไป", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtNewPassword.Text = string.Empty;
                    txtConfirmPassword.Text = string.Empty;
                    txtNewPassword.Focus();
                    return;
                }
                else
                {
                    if ((txtNewPassword.Text != txtConfirmPassword.Text))
                    {
                        MessageBox.Show("รหัสผ่านไม่ตรงกัน", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtNewPassword.Text = string.Empty;
                        txtOldPassword.Text = string.Empty;
                        txtConfirmPassword.Text = string.Empty;
                        txtOldPassword.Focus();
                        return;
                    }
                    else
                    {
                        if ((txtOldPassword.Text == txtNewPassword.Text))
                        {
                            MessageBox.Show("รหัสผ่านใหม่และเก่าเหมือนกัน", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtNewPassword.Text = string.Empty;
                            txtConfirmPassword.Text = string.Empty;
                            txtNewPassword.Focus();
                            return;
                        }
                    }
                }
                con = new OleDbConnection(cs);
                con.Open();
                var co = "Update Users set User_Password = '" + txtNewPassword.Text + "'where UserName='" + txtUserName.Text + "' and user_Password = '" + txtOldPassword.Text + "'";

                cmd = new OleDbCommand(co);
                cmd.Connection = con;
                RowsAffected = cmd.ExecuteNonQuery();
                if ((RowsAffected > 0))
                {
                    MessageBox.Show("เปลี่ยนสำเร็จ", "รหัสผ่าน", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Hide();
                    txtUserName.Text = string.Empty;
                    txtNewPassword.Text = string.Empty;
                    txtOldPassword.Text = string.Empty;
                    txtConfirmPassword.Text = string.Empty;
                    var LoginForm1 = new frmLogin();
                    LoginForm1.Show();
                    LoginForm1.txtUserName.Text = string.Empty;
                    LoginForm1.txtPassword.Text = string.Empty;
                    LoginForm1.ProgressBar1.Visible = false;
                    LoginForm1.txtUserName.Focus();
                }
                else
                {
                    MessageBox.Show("invalid user name or password", "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtUserName.Text = string.Empty;
                    txtNewPassword.Text = string.Empty;
                    txtOldPassword.Text = string.Empty;
                    txtConfirmPassword.Text = string.Empty;
                    txtUserName.Focus();
                }
                if ((con.State == ConnectionState.Open))
                {
                    con.Close();
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ล้มเหลว", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChangePassword_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            var frm = new frmLogin();
            frm.txtUserName.Text = string.Empty;
            frm.txtPassword.Text = string.Empty;
            frm.ProgressBar1.Visible = false;
            frm.txtUserName.Focus();
            frm.Show();
        }
    }
}
