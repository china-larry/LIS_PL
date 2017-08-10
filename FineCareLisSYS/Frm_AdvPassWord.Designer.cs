namespace FineCareLisSYS
{
    partial class Frm_AdvPassWord
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
            this.tb_Pw = new System.Windows.Forms.TextBox();
            this.btn_OK = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.lb_Tip = new System.Windows.Forms.Label();
            this.chk_pw = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // tb_Pw
            // 
            this.tb_Pw.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.tb_Pw.Location = new System.Drawing.Point(36, 9);
            this.tb_Pw.Name = "tb_Pw";
            this.tb_Pw.PasswordChar = '*';
            this.tb_Pw.Size = new System.Drawing.Size(186, 23);
            this.tb_Pw.TabIndex = 1;
            this.tb_Pw.Enter += new System.EventHandler(this.tb_Pw_Enter);
            // 
            // btn_OK
            // 
            this.btn_OK.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btn_OK.Location = new System.Drawing.Point(15, 57);
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.Size = new System.Drawing.Size(78, 31);
            this.btn_OK.TabIndex = 2;
            this.btn_OK.Text = "Ok";
            this.btn_OK.UseVisualStyleBackColor = true;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btn_Cancel.Location = new System.Drawing.Point(144, 57);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(78, 31);
            this.btn_Cancel.TabIndex = 3;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // lb_Tip
            // 
            this.lb_Tip.AutoSize = true;
            this.lb_Tip.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lb_Tip.ForeColor = System.Drawing.Color.Red;
            this.lb_Tip.Location = new System.Drawing.Point(55, 37);
            this.lb_Tip.Name = "lb_Tip";
            this.lb_Tip.Size = new System.Drawing.Size(140, 17);
            this.lb_Tip.TabIndex = 4;
            this.lb_Tip.Text = "Wrong password！";
            this.lb_Tip.Visible = false;
            // 
            // chk_pw
            // 
            this.chk_pw.AutoSize = true;
            this.chk_pw.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chk_pw.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chk_pw.Location = new System.Drawing.Point(15, 13);
            this.chk_pw.Name = "chk_pw";
            this.chk_pw.Size = new System.Drawing.Size(15, 14);
            this.chk_pw.TabIndex = 5;
            this.chk_pw.UseVisualStyleBackColor = true;
            this.chk_pw.CheckedChanged += new System.EventHandler(this.chk_pw_CheckedChanged);
            // 
            // Frm_AdvPassWord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(236, 92);
            this.ControlBox = false;
            this.Controls.Add(this.chk_pw);
            this.Controls.Add(this.lb_Tip);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.tb_Pw);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "Frm_AdvPassWord";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Enter password";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_Pw;
        private System.Windows.Forms.Button btn_OK;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Label lb_Tip;
        private System.Windows.Forms.CheckBox chk_pw;
    }
}