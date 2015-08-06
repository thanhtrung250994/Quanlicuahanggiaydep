namespace Quanlicuahanggiaydep
{
    partial class login
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(login));
            this.tb_usename = new System.Windows.Forms.TextBox();
            this.tb_password = new System.Windows.Forms.TextBox();
            this.lb_usename = new System.Windows.Forms.Label();
            this.lb_password = new System.Windows.Forms.Label();
            this.btn_ok_login = new System.Windows.Forms.Button();
            this.lb_login = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // tb_usename
            // 
            this.tb_usename.Location = new System.Drawing.Point(70, 20);
            this.tb_usename.Name = "tb_usename";
            this.tb_usename.Size = new System.Drawing.Size(139, 20);
            this.tb_usename.TabIndex = 0;
            // 
            // tb_password
            // 
            this.tb_password.Location = new System.Drawing.Point(70, 57);
            this.tb_password.Name = "tb_password";
            this.tb_password.PasswordChar = '*';
            this.tb_password.Size = new System.Drawing.Size(139, 20);
            this.tb_password.TabIndex = 1;
            // 
            // lb_usename
            // 
            this.lb_usename.AutoSize = true;
            this.lb_usename.Location = new System.Drawing.Point(12, 23);
            this.lb_usename.Name = "lb_usename";
            this.lb_usename.Size = new System.Drawing.Size(52, 13);
            this.lb_usename.TabIndex = 2;
            this.lb_usename.Text = "Usename";
            // 
            // lb_password
            // 
            this.lb_password.AutoSize = true;
            this.lb_password.Location = new System.Drawing.Point(12, 60);
            this.lb_password.Name = "lb_password";
            this.lb_password.Size = new System.Drawing.Size(53, 13);
            this.lb_password.TabIndex = 3;
            this.lb_password.Text = "Password";
            // 
            // btn_ok_login
            // 
            this.btn_ok_login.Location = new System.Drawing.Point(144, 83);
            this.btn_ok_login.Name = "btn_ok_login";
            this.btn_ok_login.Size = new System.Drawing.Size(65, 23);
            this.btn_ok_login.TabIndex = 4;
            this.btn_ok_login.Text = "OK";
            this.btn_ok_login.UseVisualStyleBackColor = true;
            this.btn_ok_login.Click += new System.EventHandler(this.btn_ok_login_Click);
            // 
            // lb_login
            // 
            this.lb_login.AutoSize = true;
            this.lb_login.Location = new System.Drawing.Point(12, 88);
            this.lb_login.Name = "lb_login";
            this.lb_login.Size = new System.Drawing.Size(0, 13);
            this.lb_login.TabIndex = 5;
            // 
            // login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(229, 123);
            this.Controls.Add(this.lb_login);
            this.Controls.Add(this.btn_ok_login);
            this.Controls.Add(this.lb_password);
            this.Controls.Add(this.lb_usename);
            this.Controls.Add(this.tb_password);
            this.Controls.Add(this.tb_usename);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "login";
            this.Text = "Login";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.login_FormClosing);
            this.Load += new System.EventHandler(this.Load_form);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_usename;
        private System.Windows.Forms.TextBox tb_password;
        private System.Windows.Forms.Label lb_usename;
        private System.Windows.Forms.Label lb_password;
        private System.Windows.Forms.Button btn_ok_login;
        private System.Windows.Forms.Label lb_login;
    }
}