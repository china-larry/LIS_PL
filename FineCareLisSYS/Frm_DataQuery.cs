using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace FineCareLisSYS
{
    public partial class Frm_DataQuery : Form
    {
        public Frm_DataQuery(OleDbConnection Conn)
        {
            InitializeComponent();
            mConn = Conn;
        }
        OleDbConnection mConn;
        OleDbCommand mComm;
        OleDbDataAdapter mAdpter;
        DataTable DtTable;

        private void Frm_DataQuery_Load(object sender, EventArgs e)
        {
            if (mConn.State==ConnectionState.Closed)
            {
                mConn.Open();
            }
            //填充项目名称列表
            string sql = "select ProName from Reference order by ID";
            mAdpter = new OleDbDataAdapter(sql, mConn);
            DtTable = new DataTable();
            mAdpter.Fill(DtTable);
            cmb_ItemName.Items.Clear();
            for (int i = 0; i < DtTable.Rows.Count;i++ )
            {
                cmb_ItemName.Items.Add(DtTable.Rows[i][0]);
            }
            
            //查询日期类型
            cmb_QyType.SelectedIndex = 0;
            string sItems = RWIni.ReadString("ComboBoxItems", "cmb_Doctor", "").Replace("\0", "");
            string[] sitem = sItems.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            cmb_Doctor.Items.Clear();
            for (int i = 0; i < sitem.Length; i++)
            {
                cmb_Doctor.Items.Add(sitem[i]);
            }
            string CurrPath = Directory.GetCurrentDirectory();
            sfDialog.InitialDirectory = CurrPath;
            sfDialog.DefaultExt = "*.xls";
            sfDialog.Filter = "Exel文件|*.xls";
            sfDialog.FileName = "飞测导出记录" + DateTime.Now.ToString("yyyyMMddHHmm");
        }

        private void btn_Query_Click(object sender, EventArgs e)
        {
            string sDate = DTPickerS.Value.ToString("yyyy-MM-dd") + " 00:00:00", eDate = DTPickerE.Value.ToString("yyyy-MM-dd") + " 23:59:59";
            string sSql = "select Xlh,Swatch,CheckDate,PatientName,Age,Sex,ProName,Result,Doctor,SendDateTime from Result ";
            if (cmb_QyType.SelectedIndex==0)
            {
                sSql+="where CheckDate>='" + sDate + "' and Checkdate<='" + eDate + "'";
            }
            else
            {
                sSql += "where SendDateTime>='" + sDate + "' and Checkdate<='" + eDate + "'";
            }
            if  (tb_Name.Text != "")//姓名
            {
                sSql += " and PatientName like '%" + tb_Name.Text + "%'";
            }
            if (tb_Swatch.Text != "")//样本号
            {
                sSql += " and Swatch='" + tb_Swatch.Text + "'";
            }
            if (cmb_ItemName.Text != "")//项目名称
            {
                sSql += " and ProName='" + cmb_ItemName.Text + "'";
            } if (cmb_Doctor.Text != "")//医生
            {
                sSql += " and Doctor like '%" + cmb_Doctor.Text + "%'";
            }
            sSql += " order by Xlh Desc";

            try
            {
                mAdpter = new OleDbDataAdapter(sSql, mConn);
                DtTable = new DataTable();
                mAdpter.Fill(DtTable);
                dgv_Result.Rows.Clear();

                if (DtTable.Rows.Count > 0)
                {
                    dgv_Result.Rows.Add(DtTable.Rows.Count);
                    for (int i = 0; i < DtTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < DtTable.Columns.Count; j++)
                        {
                            dgv_Result[j, i].Value = DtTable.Rows[i][j];

                        }
                    }
                }
            } 
            catch (System.Exception ex) {}
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            string fileName = "";
            if (sfDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = sfDialog.FileName;
                //保存到Excel中
                Excel.Application xlsApp = new Excel.Application();
                Excel.Workbook wb1 = xlsApp.Workbooks.Add(true); //xlsApp.Workbooks.Add(true)表示创建一个xls
                Excel.Worksheet ws1 = wb1.Worksheets[1] as Excel.Worksheet;

                for (int j = 0; j < dgv_Result.ColumnCount; j++)
                {
                    ws1.Cells[1, j + 1] = dgv_Result.Columns[j].HeaderText;
                }
                for (int i = 0; i < DtTable.Rows.Count; i++)
                {
                    for (int j = 0; j < DtTable.Columns.Count; j++)
                    {
                        ws1.Cells[i + 2, j + 1] = "'" + DtTable.Rows[i][j].ToString();//以文本格式保存

                    }
                }
                ws1.Columns.AutoFit();//自动适应列宽
                wb1.SaveCopyAs(fileName);
                wb1.Close(false, null, null);
                xlsApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
                MessageBox.Show("导出成功!", "提示");
            } 
        }
    }
}
