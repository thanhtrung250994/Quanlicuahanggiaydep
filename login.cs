using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quanlicuahanggiaydep
{
    public partial class login : Form
    {
        public string usename = null;
        public string password = null;
        OleDbConnection con = new OleDbConnection();
        public login()
        {
            InitializeComponent();
        }

        private void btn_ok_login_Click(object sender, EventArgs e)
        {   Main_form main_form = new Main_form();
            OleDbCommand command = new OleDbCommand();
            command.Connection = con;
            //command.CommandText = "select * from Nhan_vien where Ma_nv= '" + tb_usename.Text + "' and Password='" + tb_password.Text + "'";
            command.CommandText = "select * from Nhan_vien where Ma_nv= @username and Passwords= @pass";
            command.Parameters.AddWithValue("@username", tb_usename.Text);
            command.Parameters.AddWithValue("@pass", tb_password.Text);

            try
            {
                OleDbDataReader reader = command.ExecuteReader();
                int count = 0;
                while (reader.Read())
                {
                    count++;
                    main_form.doanh_thu_cua_nv_dang_dang_nhap = reader.GetInt32(3);
                    main_form.Uu_tien = reader.GetInt32(8);
                }
                if (count == 1)
                {
                    
                    this.Hide();
                    main_form.Ma_nv_login = tb_usename.Text;
                    main_form.ShowDialog();
                    con.Close();
                }
                else
                { lb_login.Text = "Use hoặc pass sai!"; }
            }
            catch { }
        }

        private void Load_form(object sender, EventArgs e)
        {
            try
            {
                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Quan_ly_ban_hang_giay_dep.accdb";
                con.Open();
                Debug.WriteLine("Connect successful");
                //con.Close();
            }
            catch { Debug.WriteLine("Connect error"); }
        }

        private void login_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = false;
        }
    }
}
