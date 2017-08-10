using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FineCareLisSYS
{
    public partial class Frm_AdvPassWord : Form
    {
        public Frm_AdvPassWord()
        {
            InitializeComponent();
        }
        public bool rightPw = false;
        public int PwdLevel = 0;
        private void btn_OK_Click(object sender, EventArgs e)
        {
            if (tb_Pw.Text=="wf2016")
            {
                rightPw = true;
                PwdLevel = 2;
                this.Close();
            }
            else if (tb_Pw.Text == "wf2017")
            {
                rightPw = true;
                PwdLevel = 1;
                this.Close();
            }
            else
            {
                lb_Tip.Visible = true;
                tb_Pw.Text = ""; rightPw = false;
            }
        }
        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();  rightPw = false;
        }
        private void tb_Pw_Enter(object sender, EventArgs e)
        {
            lb_Tip.Visible = false;
        }

        private void chk_pw_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_pw.Checked)
            {
                tb_Pw.PasswordChar = '\0';
            }
            else
                tb_Pw.PasswordChar = '*';
        }   
    }
}
