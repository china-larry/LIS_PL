using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FineCareLisSYS
{
    public partial class mInputBox : Form
    {

        public string OutText="";
        public mInputBox()
        {
            InitializeComponent();
        }
        private void mInputBox_Load(object sender, EventArgs e)
        {
            tb_Input.Focus();
            tb_Input.SelectAll();
        }
        private void btn_OK_Click(object sender, EventArgs e)
        {
            OutText = tb_Input.Text;
            this.Close();
        }

        private void mInputBox_Activated(object sender, EventArgs e)
        {
            tb_Input.Focus();
            tb_Input.SelectAll();
        }

        

        
    }
}
;