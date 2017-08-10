using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Data.OleDb;
using System.Timers;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;



namespace FineCareLisSYS
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            FocusControls = new List<Control>();
            FocusControls.Add(tb_Name);
            FocusControls.Add(cmb_Sex);
            FocusControls.Add(tb_Age);
            FocusControls.Add(tb_HosptNo);
            FocusControls.Add(tb_BedNo);
            FocusControls.Add(cmb_Office);
            FocusControls.Add(cmb_SampleType);
            FocusControls.Add(cmb_Doctor);
            FocusControls.Add(cmb_Tester);
            FocusControls.Add(cmb_Audit);
            FocusControls.Add(tb_Diagnosis);

            for (int i = 0; i < dgv_Result.ColumnCount; i++)//默认设置成不可编辑
            {
                dgv_Result.Columns[i].ReadOnly = true;
            }

            ThreadPool.SetMaxThreads(30, 30);
            Control.CheckForIllegalCrossThreadCalls = false;
        }
        string spath = Directory.GetCurrentDirectory();
        OleDbConnection conn;
        OleDbDataAdapter dtAdpterResult;
        OleDbCommand oleInsertData;
        string strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Directory.GetCurrentDirectory() + "\\WF2010.mdb;Jet OLEDB:Database Password=wf2016";//Persist Security Info=False;
        string strSqlresult = "select Xlh,Swatch,ProName,Result,DW_Unit,CheckDate,PatientName,Sex,Age,HosNo,BedNo," +
                    "HOffice,Sample,Doctor,Scrutator,Audit,ClinicalDiagnosis,SendDateTime,DW_Reference,DW_Description from Result order by Xlh";

        private List<Control> FocusControls;

        WFReport wfReport1 = new WFReport();

        private void Form1_Load(object sender, EventArgs e)
        {
            //获取数据源
            conn = new OleDbConnection(strConn);
            conn.Open();
            dtAdpterResult = new OleDbDataAdapter(strSqlresult, conn);

            //加载水晶报表 
            wfReport1.Load(spath + "\\WFReport.rpt");

            //加载下拉列表项
            cmb_Sex.SelectedIndex = 0; cmb_SampleType.SelectedIndex = 0;

            cmb_Doctor.Items.Clear(); 
            cmb_Office.Items.Clear(); 
            cmb_Tester.Items.Clear(); 
            cmb_Audit.Items.Clear();
           
            string sItems= RWIni.ReadString("ComboBoxItems", "cmb_Office", "").Replace("\0","");
            string[] sitem = sItems.Split(new char[] { ',' },StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < sitem.Length; i++)
            {
                cmb_Office.Items.Add(sitem[i]);
            }
                cmb_Office.Text = string.Empty;
            sItems = RWIni.ReadString("ComboBoxItems", "cmb_Retester", "").Replace("\0", "");
            sitem = sItems.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < sitem.Length; i++)
            {
                cmb_Tester.Items.Add(sitem[i]);
            }
                cmb_Tester.Text = string.Empty;
            sItems = RWIni.ReadString("ComboBoxItems", "cmb_Audit", "").Replace("\0", "");
            sitem = sItems.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < sitem.Length; i++)
            {
                cmb_Audit.Items.Add(sitem[i]);
            }
            cmb_Audit.Text = string.Empty;
            sItems = RWIni.ReadString("ComboBoxItems", "cmb_Doctor", "").Replace("\0", "");
            sitem = sItems.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < sitem.Length; i++)
            {
                cmb_Doctor.Items.Add(sitem[i]);
            }
                cmb_Doctor.Text=string.Empty;

            btn_StartRecv.Enabled = true;
            btn_StopRecv.Enabled = false;
            btn_StartRecv_Click(sender, e);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr=MessageBox.Show("确认退出软件?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.Yes)
            {
                if (UdpRecv != null)
                {
                    try
                    {
                        UdpRecv.Close();
                    }
                    catch (System.Net.Sockets.SocketException ex)
                    {
                        UdpRecv = null;
                    }
                }
                dtAdpterResult.Dispose();               
                conn.Close();
            }
            else
            {
                e.Cancel = true;
            }
        }

        //选择结果
        private void dgv_Result_SelectionChanged(object sender, EventArgs e)
        {
            if (dgv_Result.CurrentRow == null)
            {
                return;
            }
            int crrRow = dgv_Result.CurrentRow.Index;
            if (dgv_Result[1, crrRow].Value == null)
            {
                return;
            }
            string sResult = "", sSwatch = "", sUnit = "", sProName = "", sCheckDate = "";
            try
            {                
                sProName = dgv_Result[col_ProName.Index, crrRow].Value.ToString();
                sSwatch = dgv_Result[col_Swatch.Index, crrRow].Value.ToString();
                sResult = dgv_Result[col_Result.Index, crrRow].Value.ToString();
                sUnit = dgv_Result[col_Unit.Index, crrRow].Value.ToString();
                sCheckDate = dgv_Result[col_CheckTime.Index, crrRow].Value.ToString();
            }
            catch (System.Exception ex)
            {
                return;
            }
            richTb_Info.Text = "检测结果：" + "\r\n" + "项目名称：" + sProName + "\r\n" + "样本号：" + sSwatch
                + "\r\n" + "结果：" + sResult + sUnit + "\r\n" + "检测时间：" + sCheckDate;
            //更新文本框
                tb_Name.Text = dgv_Result[col_Name.Index, crrRow].Value.ToString();
                cmb_Sex.Text = dgv_Result[col_Sex.Index, crrRow].Value.ToString();
                tb_Age.Text = dgv_Result[col_Age.Index, crrRow].Value.ToString();
                tb_HosptNo.Text = dgv_Result[col_HosNo.Index, crrRow].Value.ToString();
                tb_BedNo.Text = dgv_Result[col_BedNo.Index, crrRow].Value.ToString();
                cmb_Office.Text = dgv_Result[col_Office.Index, crrRow].Value.ToString();
                cmb_SampleType.Text = dgv_Result[col_Sample.Index, crrRow].Value.ToString();
                cmb_Doctor.Text = dgv_Result[col_Doc.Index, crrRow].Value.ToString();
                cmb_Tester.Text = dgv_Result[col_Retester.Index, crrRow].Value.ToString();
                cmb_Audit.Text = dgv_Result[col_Audit.Index, crrRow].Value.ToString();
            //16列是临床诊断
                tb_Diagnosis.Text = dgv_Result[col_Dialgnosis.Index, crrRow].Value.ToString();
        }

        private void btn_StopRecv_Click(object sender, EventArgs e)
        {
            try
            {
                UdpRecv.Close();
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                UdpRecv = null;
            }
        }

        //接收处理线程
        private void btn_StartRecv_Click(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(new WaitCallback(ReceviceData));
        }

        private IPEndPoint IPEnd;
        private UdpClient UdpRecv;
        private void ReceviceData(object ob) 
        {
            btn_StartRecv.Enabled = false;
            btn_StopRecv.Enabled = true;
            IPEnd = new IPEndPoint(IPAddress.Any, 0);//IP终端, 接收IPend,任意端口
            UdpRecv = new UdpClient(int.Parse(tb_Port.Text));//本机端口
            while (true)
            {
                try
                {
                    byte[] recvBuff = UdpRecv.Receive(ref IPEnd);
                    if (recvBuff.Length > 0)
                    {
                        string sRecv = Encoding.UTF8.GetString(recvBuff);
                        //处理数据
                        sRecv = sRecv.Replace("FF&", "").Replace("&EE", "");//去掉包头包尾
                        string[] sSec = sRecv.Split(new char[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                        if (sSec.Length == 5 && sRecv.IndexOf("CRP") != 0)
                        {
                            string[] sCRP = new string[4];
                            string[] sHsCRP = new string[4];
                            sCRP[0] = sSec[0];
                            sHsCRP[0] = sSec[0];
                            sCRP[1] = sSec[1];
                            sHsCRP[1] = sSec[2];
                            sCRP[2] = sSec[3];
                            sHsCRP[2] = sSec[3];
                            sCRP[3] = "CRP";
                            sHsCRP[3] = "HsCRP";
                            InsertNewRow(sHsCRP);
                            InsertNewRow(sCRP);
                        }
                        else
                        {
                            InsertNewRow(sSec);
                        }
                    }
                    Thread.Sleep(500);
                }
                catch(System.Net.Sockets.SocketException ex)
                {
                    break;
                }
            }
            UdpRecv = null;
            btn_StartRecv.Enabled = true;
            btn_StopRecv.Enabled = false;
        }

        private void InsertNewRow(string[] sSec)
        {
            string[] strInit = new string[dgv_Result.ColumnCount];
            for (int i = 0; i < strInit.Length; i++)
            strInit[i] = "";
            strInit[1] = sSec[0];//样本号
            strInit[3] = sSec[1];//结果
            strInit[5] = sSec[2];//检测时间
            try
            {
                //显示名称获取
                OleDbCommand oleComm = new OleDbCommand("select ProName from Reference where ProName='" + sSec[3] + "'", conn);
                string sPrintName = oleComm.ExecuteScalar().ToString();
                strInit[2] = sPrintName;
                //单位获取
                oleComm = new OleDbCommand("select mUnit from DW_Unit where ProName='" + sSec[3] + "'", conn);
                string sUnit = oleComm.ExecuteScalar().ToString();
                strInit[4] = sUnit;
                //参考范围
                oleComm.CommandText = "select mReference from Reference where ProName='" + sSec[3] + "'";
                string sRef = oleComm.ExecuteScalar().ToString();
                strInit[strInit.Length - 2] = sRef;
                //参考范围说明
                oleComm.CommandText = "select mDescription from Reference where ProName='" + sSec[3] + "'";
                string sDescrip = oleComm.ExecuteScalar().ToString();
                strInit[strInit.Length - 1] = sDescrip;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("未查到该项目对应录入的单位或参考范围,可以先到【单位设置】录入!");
            }
            //接收时间
            strInit[strInit.Length - 3] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            dgv_Result.Rows.Insert(0, strInit);//添加一行

            //在数据库添加数据
            OleDbCommand oleMaxXlh = new OleDbCommand("select max(Xlh) from Result", conn);
            int Xlh = Convert.ToInt32(oleMaxXlh.ExecuteScalar().ToString());
            Xlh++;
            string ResultStirng = string.Format(@"insert into Result(Xlh, Swatch, Result, CheckDate, ProName, DW_Unit, DW_Reference, DW_Description, SendDateTime)values({0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')",
                Xlh, strInit[1], strInit[3], strInit[5], strInit[2], strInit[4], strInit[strInit.Length - 2], strInit[strInit.Length - 1], strInit[strInit.Length - 3]);
            using (oleInsertData = new OleDbCommand(ResultStirng, conn))
            {
                int i =oleInsertData.ExecuteNonQuery();
                oleInsertData.Dispose();
            }
        }
        //单位设置
        private void btn_UnitSet_Click(object sender, EventArgs e)
        {
            FrmSetting frmSet = new FrmSetting();
            frmSet.ShowDialog(this);
        }

        //保存按钮
        private void btn_Save_Click(object sender, EventArgs e)
        {
            if (dgv_Result.SelectedRows.Count > 0)
            {
                int icrrRow = dgv_Result.SelectedRows[dgv_Result.SelectedRows.Count - 1].Index;  //当前行
                try
                {
                    string UpdateString = string.Format(@"update Result set PatientName='{0}', Sex='{1}' , Age='{2}', HosNo='{3}', BedNo='{4}', HOffice='{5}', Sample='{6}', Doctor='{7}', ClinicalDiagnosis='{8}', Scrutator='{9}', Audit='{10}' where Xlh={11}",
                    tb_Name.Text, cmb_Sex.Text, tb_Age.Text, tb_HosptNo.Text, tb_BedNo.Text, cmb_Office.Text, cmb_SampleType.Text, cmb_Doctor.Text, tb_Diagnosis.Text, cmb_Tester.Text, cmb_Audit.Text, dgv_Result[col_xlh.Index, icrrRow].Value);
                    using (oleInsertData = new OleDbCommand(UpdateString, conn))
                    {
                        int i = oleInsertData.ExecuteNonQuery();
                        oleInsertData.Dispose();
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("信息录入失败.","错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                dgv_Result[col_Name.Index, icrrRow].Value = tb_Name.Text;
                dgv_Result[col_Sex.Index, icrrRow].Value = cmb_Sex.Text;
                dgv_Result[col_Age.Index, icrrRow].Value = tb_Age.Text;
                dgv_Result[col_HosNo.Index, icrrRow].Value = tb_HosptNo.Text;
                dgv_Result[col_BedNo.Index, icrrRow].Value = tb_BedNo.Text;
                dgv_Result[col_Office.Index, icrrRow].Value = cmb_Office.Text;
                dgv_Result[col_Sample.Index, icrrRow].Value = cmb_SampleType.Text;//样本类型
                dgv_Result[col_Doc.Index, icrrRow].Value = cmb_Doctor.Text;
                dgv_Result[col_Dialgnosis.Index, icrrRow].Value = tb_Diagnosis.Text;//临床诊断               
                dgv_Result[col_Retester.Index, icrrRow].Value = cmb_Tester.Text;
                dgv_Result[col_Audit.Index, icrrRow].Value = cmb_Audit.Text;

            }
            //保存下拉列表选项cmb_Office,cmb_Doctor,cmb_Retester,cmb_Audit
            string sItems = "";
            foreach (object itm in cmb_Office.Items)
            {
                sItems += itm.ToString() + ",";
            }
            if (sItems.IndexOf(cmb_Office.Text)==-1)//新增项
            {
                cmb_Office.Items.Insert(0, cmb_Office.Text);//添加到列表
                sItems = cmb_Office.Text + "," + sItems;
            }
            RWIni.WriteString("ComboBoxItems", "cmb_Office", sItems.Substring(0,sItems.Length-1));//去掉最后一个逗号
            sItems = "";
            foreach (object itm in cmb_Doctor.Items)
                sItems += itm.ToString() + ",";
            if (sItems.IndexOf(cmb_Doctor.Text) == -1)//新增项
            {
                cmb_Doctor.Items.Insert(0, cmb_Doctor.Text);
                sItems = cmb_Doctor.Text + "," + sItems;
            }
            RWIni.WriteString("ComboBoxItems", "cmb_Doctor", sItems.Substring(0, sItems.Length - 1));
            sItems = "";
            foreach (object itm in cmb_Tester.Items)
                sItems += itm.ToString() + ",";
            if (sItems.IndexOf(cmb_Tester.Text) == -1)//新增项
            {
                cmb_Tester.Items.Insert(0, cmb_Tester.Text);
                sItems = cmb_Tester.Text + "," + sItems;
            }
            RWIni.WriteString("ComboBoxItems", "cmb_Retester", sItems.Substring(0, sItems.Length - 1));
            sItems = "";
            foreach (object itm in cmb_Audit.Items)
                sItems += itm.ToString() + ",";
            if (sItems.IndexOf(cmb_Audit.Text) == -1)//新增项
            {
                cmb_Audit.Items.Insert(0, cmb_Audit.Text);
                sItems = cmb_Audit.Text + "," + sItems;
            }
            RWIni.WriteString("ComboBoxItems", "cmb_Audit", sItems.Substring(0, sItems.Length - 1));

            statusStrip1.Items[1].Text = "提示信息: 保存当前行成功";

        }
        private int GetMaxID()
        {
            int iret = 0 ;
            OleDbCommand cmd = new OleDbCommand("select max(Xlh) from Result", conn);
            try
            {
                iret=int.Parse(cmd.ExecuteScalar().ToString());
            }
            catch (System.Exception ex)
            {
            }
            return iret;
        }

        //查询按钮
        private void btn_Query_Click(object sender, EventArgs e)
        {
            string Sql;
            if (chk_QueryAll.Checked)
            {
                Sql = "select Xlh,Swatch,ProName,Result,DW_Unit,CheckDate,PatientName,Sex,Age,HosNo,BedNo," +
                      "HOffice,Sample,Doctor,Scrutator,Audit,ClinicalDiagnosis,SendDateTime,DW_Reference,DW_Description from Result order by Xlh Desc";
            }
            else
            {
                Sql = "select Xlh,Swatch,ProName,Result,DW_Unit,CheckDate,PatientName,Sex,Age,HosNo,BedNo," +
                      "HOffice,Sample,Doctor,Scrutator,Audit,ClinicalDiagnosis,SendDateTime,DW_Reference,DW_Description from Result where CheckDate like'" + DTPicker1.Value.ToString("yyyy-MM-dd") + "%' order by Xlh Desc";
            }

            OleDbDataAdapter mAdpt = new OleDbDataAdapter(Sql, conn);
            DataTable mtb = new DataTable();
            mAdpt.Fill(mtb);
            dgv_Result.Rows.Clear();

            if (mtb.Rows.Count > 0)
            {
                dgv_Result.Rows.Add(mtb.Rows.Count);
                for (int i = 0; i < mtb.Rows.Count; i++)
                {
                    for (int j = 0; j < mtb.Columns.Count; j++)
                    {
                        dgv_Result[j, i].Value = mtb.Rows[i][j];

                    }
                }
            }
        }

        //打印
        private void btn_Print_Click(object sender, EventArgs e)
        {

            if (dgv_Result.SelectedRows.Count==0)
            {
                MessageBox.Show("未选择打印内容.", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dgv_Result.SelectedRows.Count > 6)
            {
                MessageBox.Show("打印内容最多选择6项.", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

                ReportParmeters RP = new ReportParmeters(dgv_Result);
                ResulteTableValue RV = RP.GetReportInformatiom();
                if (RV != null)
                {
                    try
                    {
                    wfReport1.SetParameterValue("医院名称", RWIni.ReadString("SystemSetting", "HospitalName", "请到Config.ini文件的HospitalName字段中修改医院名称!"));
                    wfReport1.SetParameterValue("姓名", RV.Name);
                    wfReport1.SetParameterValue("性别", RV.Sex);
                    wfReport1.SetParameterValue("年龄", RV.Age);
                    wfReport1.SetParameterValue("住院号", RV.HosNo);
                    wfReport1.SetParameterValue("床号", RV.BedNo);
                    wfReport1.SetParameterValue("科室", RV.Office);
                    wfReport1.SetParameterValue("医生", RV.Doc);
                    wfReport1.SetParameterValue("操作者", RV.Retester);
                    wfReport1.SetParameterValue("接收时间", RV.RecvTime);

                    for (int i = 0; i <6; i++)
                    {
                        if (i < RP.SelectionLength)
                        {
                            wfReport1.SetParameterValue("项目名称1" + (i + 1), RP.ReportValue[i].ProName);
                            wfReport1.SetParameterValue("结果1" + (i + 1), RP.ReportValue[i].Result);
                            wfReport1.SetParameterValue("单位1" + (i + 1), RP.ReportValue[i].Unit);
                            wfReport1.SetParameterValue("实验参考范围1" + (i + 1), RP.ReportValue[i].Reference);
                            wfReport1.SetParameterValue("实验参考范围说明1" + (i + 1), RP.ReportValue[i].Description);
                        }
                        else
                        {
                            wfReport1.SetParameterValue("项目名称1" + (i + 1), string.Empty);
                            wfReport1.SetParameterValue("结果1" + (i + 1), string.Empty);
                            wfReport1.SetParameterValue("单位1" + (i + 1), string.Empty);
                            wfReport1.SetParameterValue("实验参考范围1" + (i + 1), string.Empty);
                            wfReport1.SetParameterValue("实验参考范围说明1" + (i + 1), string.Empty);
                        }
                    }

                    PrintForm frm_report = new PrintForm();
                    frm_report.crystalReportViewer1.ReportSource = wfReport1;
                    if (chk_PrintView.Checked == true)
                        frm_report.Show();
                    else
                        frm_report.crystalReportViewer1.PrintReport();
                    }

                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString()); 
                    }
                }
        }

        private void btn_Statistic_Click(object sender, EventArgs e)
        {
            Frm_DataQuery frmDataQ = new Frm_DataQuery(conn);
            frmDataQ.ShowDialog(this);
        }

        private void btn_Abaut_Click(object sender, EventArgs e)
        {
            About frmabout = new About();
            frmabout.ShowDialog(this);
        }

        //勾选启用编辑
        private void chk_Edit_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_Edit.Checked)//勾上可编辑
            {
                Frm_AdvPassWord frmPw = new Frm_AdvPassWord();
                frmPw.ShowDialog(this);
                if (frmPw.rightPw)
                {
                    if (frmPw.PwdLevel == 1)
                    {
                        for (int i = 1; i < dgv_Result.ColumnCount; i++)
                        {
                            dgv_Result.Columns[i].ReadOnly = false;
                        }
                        dgv_Result.AllowUserToDeleteRows = true;
                        lb_EditTip.Visible = true;
                        btn_DelCurr.Visible = true;
                    }
                    else
                    {
                        chk_Edit.Checked = false;
                        MessageBox.Show("密码等级错误.");
                    }
                }
                else
                    chk_Edit.Checked = false;
            }
            else
            {
                for (int i = 1; i < dgv_Result.ColumnCount; i++)
                {
                    dgv_Result.Columns[i].ReadOnly = true;//不可编辑
                }
                dgv_Result.AllowUserToDeleteRows = false;//不可删除
                lb_EditTip.Visible = false;
                btn_DelCurr.Visible = false;
            }
        }
        //删除
        private void btn_DelCurr_Click(object sender, EventArgs e)
        {
            if (dgv_Result.CurrentRow==null)
            {
                return;
            }
            if (MessageBox.Show("确定删除当前行?", "系统提示") == DialogResult.OK)
            {
                try
                {
                    DataGridViewRow currRow = dgv_Result.CurrentRow;
                    dgv_Result.Rows.Remove(currRow);
                    //删除数据库中对应行
                    string xlh=currRow.Cells[0].Value.ToString();
                    string sql = "delete from Result where Xlh=" + (xlh == "" ? "0" : xlh);
                    OleDbCommand comm = new OleDbCommand(sql, conn);
                    comm.CommandType = CommandType.Text;
                    comm.ExecuteNonQuery();
                    statusStrip1.Items[1].Text = "提示信息: 删除成功.";
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }


        //管理ComboBox的下拉列表
        ComboBox cmb_curr = new ComboBox();
        private void cmb_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                cmb_curr = ((ComboBox)sender);
                cmb_curr.Focus();
            }
        }

        private void Menu_ItemDel_Click(object sender, EventArgs e)//删除
        {
            cmb_curr.Items.Remove(cmb_curr.SelectedItem);
            string sItems = "";
            foreach (object itm in cmb_curr.Items)
            {
                sItems += itm.ToString() + ",";
            }
            RWIni.WriteString("ComboBoxItems", cmb_curr.Name, sItems.Substring(0, sItems.Length - 1));//去掉最后一个逗号
            if (cmb_curr.Items.Count > 0)
                cmb_curr.SelectedIndex = 0;
            if (cmb_curr.Items.Count == 0)
                cmb_curr.Text = "";
        }

        private void menu_IemAdd_Click(object sender, EventArgs e)//新增
        {
            mInputBox mInput = new mInputBox();
            mInput.ShowDialog(this);
            if (mInput.OutText=="")
            {
                return;
            }
            string addstr=mInput.OutText, sItems = "";
            if (cmb_curr.Items.IndexOf(addstr) == -1)
            {
                cmb_curr.Items.Insert(0, addstr);
            }
            foreach (object itm in cmb_curr.Items)
            {
                sItems += itm.ToString() + ",";
            }
            RWIni.WriteString("ComboBoxItems", cmb_curr.Name, sItems.Substring(0, sItems.Length - 1));//去掉最后一个逗号
            cmb_curr.SelectedIndex = 0;
        }

        private void chk_QueryAll_CheckedChanged(object sender, EventArgs e)
        {
            DTPicker1.Enabled =! chk_QueryAll.Checked;
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)  
        {  
            if ( e.KeyData == Keys.Enter )  
            {
                Control Ctl = MainContainer.ActiveControl;
                int ContrlIndex = FindControl(Ctl, FocusControls);
                if (ContrlIndex != -1 && ContrlIndex < FocusControls.Count - 1)
                {
                    ContrlIndex++;
                }
                else ContrlIndex = 0;
                FocusControls[ContrlIndex].Focus();
            }  
        }

        private int FindControl(Control Ctl, List<Control> FCtls)
        {
            for (int i = 0; i < FCtls.Count; i++)
            {
                if (FocusControls[i].Name == Ctl.Name)
                    return i;
            }
            return -1;
        }
    }
}
