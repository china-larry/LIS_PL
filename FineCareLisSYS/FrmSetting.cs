using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace FineCareLisSYS
{
    public partial class FrmSetting : Form
    {
        public FrmSetting()
        {
            InitializeComponent();
        }
        private void btn_ExitSetting_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        OleDbConnection conn;
        OleDbCommand comm;
        private void FrmSetting_Load(object sender, EventArgs e)
        {
            try
            {
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Directory.GetCurrentDirectory() + "\\WF2010.mdb;Jet OLEDB:Database Password=wf2017 ";
                string strSql = "select Reference.ProName,DW_Unit.mUnit,Reference.mReference, Reference.mPrintName, Reference.mDescription from DW_Unit , Reference where DW_Unit.ProName=Reference.ProName order by Reference.ID";
                conn = new OleDbConnection(strConn);
                conn.Open();
                comm = new OleDbCommand(strSql, conn);
                OleDbDataReader dr = comm.ExecuteReader();
                int xh = 1;
                while (dr.Read())
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = (xh++).ToString();
                    lvi.SubItems.Add(dr.GetString(0));//项目名称
                    try
                    {
                        lvi.SubItems.Add(dr.GetString(1));//单位
                    }
                    catch (System.Exception ex)
                    {
                        lvi.SubItems.Add("");
                    }
                    try
                    {
                        lvi.SubItems.Add(dr.GetString(2));//范围
                    }
                    catch (System.Exception ex)
                    {
                        lvi.SubItems.Add("");
                    }
                    try
                    {
                        lvi.SubItems.Add(dr.GetString(3));//显示名称
                    }
                    catch (System.Exception ex)
                    {
                        lvi.SubItems.Add("");
                    }
                    try
                    {
                        lvi.SubItems.Add(dr.GetString(4));
                    }
                    catch (System.Exception ex)
                    {
                        lvi.SubItems.Add("");
                    }
                    lstV1.Items.Add(lvi);

                }
                conn.Close();
            }
            catch (System.Exception ex)
            {

            }
        }

        private void lstV1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstV1.SelectedItems.Count > 0)
            {
                ListViewItem lvi = lstV1.SelectedItems[0];
                tb_Name.Text = lvi.SubItems[1].Text;
                tb_Unit.Text = lvi.SubItems[2].Text;
                tb_Reference.Text = lvi.SubItems[3].Text;
                tb_PrintName.Text = lvi.SubItems[4].Text;
                tb_Description.Text = lvi.SubItems[5].Text;
                tb_AddItem.Text = "";
                tb_Name.Refresh();
            }
        }

        private void chk_reference_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_reference.Checked)
            {
                Frm_AdvPassWord frmPw = new Frm_AdvPassWord();
                frmPw.ShowDialog(this);
                if (frmPw.rightPw)
                    if (frmPw.PwdLevel == 2)//二级密码
                    {
                        tb_Reference.Enabled = true;
                        tb_Description.Enabled = true;
                    }
                    else
                    {
                        chk_reference.Checked = false;
                        MessageBox.Show("密码等级错误!");
                    }
                else
                    chk_reference.Checked = false;
            }
            else
            {
                tb_Reference.Enabled = false;
                tb_Description.Enabled = false;
            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                string sql = "";
                ListViewItem lvi = lstV1.SelectedItems[0];//选中行的第一行
                if (lvi.SubItems[2].Text != tb_Unit.Text)
                {
                    sql = "update DW_Unit set mUnit='" + tb_Unit.Text + "' where ProName='" + tb_Name.Text + "'";
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                    lvi.SubItems[2].Text = tb_Unit.Text;
                }
                if (lvi.SubItems[3].Text != tb_Reference.Text)
                {
                    sql = "update Reference set mReference='" + tb_Reference.Text + "' where ProName='" + tb_Name.Text + "'";
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                    lvi.SubItems[3].Text = tb_Reference.Text;
                }
                if (lvi.SubItems[4].Text != tb_PrintName.Text)
                {
                    sql = "update Reference set mPrintName='" + tb_PrintName.Text + "' where ProName='" + tb_Name.Text + "'";
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                    lvi.SubItems[4].Text = tb_PrintName.Text;
                }
                if (lvi.SubItems[5].Text != tb_Description.Text)
                {
                    sql = "update Reference set mDescription='" + tb_Description.Text + "' where ProName='" + tb_Name.Text + "'";
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                    lvi.SubItems[5].Text = tb_Description.Text;
                }
                conn.Close();
            }
            catch (System.Exception ex)
            {
                return;
            }
            MessageBox.Show("保存成功,将对新收的记录生效! ", "系统提示");
        }

        private void btn_AddItem_Click(object sender, EventArgs e)
        {
            if (tb_AddItem.Text == "")
            {
                MessageBox.Show("请输入要新增的项目名称!");
                return;
            }
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            try
            {
                comm.CommandText = "select * from DW_Unit where ProName='" + tb_AddItem.Text + "'";
                OleDbDataReader dr = comm.ExecuteReader();
                if (!dr.HasRows)
                {
                    dr.Close();
                    string sql = "insert into DW_Unit(ProName) values('" + tb_AddItem.Text + "')";
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                }
                dr.Close();
                comm.CommandText = "select * from Reference where ProName='" + tb_AddItem.Text + "'";
                dr = comm.ExecuteReader();
                if (!dr.HasRows)
                {
                    dr.Close();
                    string sql = "insert into Reference(ProName) values('" + tb_AddItem.Text + "')";
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                    //添加到控件
                    int xh = lstV1.Items.Count;
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = (xh + 1).ToString();
                    lvi.SubItems.Add(tb_AddItem.Text);//项目名称
                    for (int i = 0; i < 4; i++)
                    {
                        lvi.SubItems.Add("");
                    }
                    lstV1.Items.Add(lvi);
                    lstV1.Focus();
                    lstV1.Items[lstV1.Items.Count - 1].Selected = true;
                }
                else
                {
                    MessageBox.Show("已存在该项目!");
                }
                dr.Close();
            }
            catch (System.Exception ex)
            {

            }
            conn.Close();
        }
    }
}
