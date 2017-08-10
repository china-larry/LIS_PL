using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FineCareLisSYS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“wFDataSet1.Result”中。您可以根据需要移动或移除它。
            //this.resultTableAdapter2.Fill(this.wFDataSet1.Result);
            // TODO: 这行代码将数据加载到表“wFDataSet.Result”中。您可以根据需要移动或移除它。
            this.resultTableAdapter1.Fill(this.WFDataSet.Result);
            // TODO: 这行代码将数据加载到表“wFDataSet.Result”中。您可以根据需要移动或移除它。
            //this.resultTableAdapter.Fill(this.wFDataSet.Result);

        }
    }
}
