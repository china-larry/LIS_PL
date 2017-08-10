using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace FineCareLisSYS
{
    public class ResulteTableValue 
    {
        public string Xlh { get; private set; }                         //序号
        public string Swatch { get; private set; }                  //样本号       
        public string Sample { get; private set; }                  //样本类型
        public string ProName { get; private set; }               //项目名
        public string Result { get; private set; }                    //结果
        public string Unit { get; private set; }                       //单位
        public string Name { get; private set; }                     //姓名
        public string Sex { get; private set; }                         //性别
        public string Age { get; private set; }                        //年龄
        public string HosNo { get; private set; }                    //住院号
        public string BedNo { get; private set; }                    //床号
        public string Office { get; private set; }                     //科室
        public string RecvTime { get; private set; }                //接收时间
        public string CheckTime { get; private set; }                //检测时间
        public string Reference { get; private set; }               //实验参考范围
        public string Description { get; private set; }             //实验参考范围说明
        public string Doc { get; private set; }                        //医生
        public string Retester { get; private set; }                  //操作者

        public ResulteTableValue(DataGridViewRow GR)
        {
            Xlh = GR.Cells[0].Value.ToString();
            Swatch = GR.Cells[1].Value.ToString();
            ProName = GR.Cells[2].Value.ToString();
            Result = GR.Cells[3].Value.ToString();
            Unit = GR.Cells[4].Value.ToString();
            CheckTime = GR.Cells[5].Value.ToString();
            Name = GR.Cells[6].Value.ToString();
            Sex = GR.Cells[7].Value.ToString();
            Age = GR.Cells[8].Value.ToString();
            HosNo = GR.Cells[9].Value.ToString();
            BedNo = GR.Cells[10].Value.ToString();
            Office = GR.Cells[11].Value.ToString();
            RecvTime = GR.Cells[17].Value.ToString();
            Reference = GR.Cells[18].Value.ToString();
            Description = GR.Cells[19].Value.ToString();
            Sample = GR.Cells[12].Value.ToString();
            Doc = GR.Cells[13].Value.ToString();
            Retester = GR.Cells[14].Value.ToString();
        }
    }

    public class ReportParmeters
    {
        private ResulteTableValue[ ] SeletResults;
        public int SelectionLength { get; private set; }
        public ResulteTableValue[ ] ReportValue { get { return SeletResults; } }

        public ReportParmeters(DataGridView GV) 
        {
            SelectionLength = GV.SelectedRows.Count;
            SeletResults = new ResulteTableValue[SelectionLength];
            for (int i = 0; i < SelectionLength; i++)
            {
                SeletResults[i]=new ResulteTableValue(GV.SelectedRows[i]);
            }
        }

        public ResulteTableValue GetReportInformatiom()
        {
            List<int> EmptyIndex = new List<int>();
            List<int> NotEmptyIndex = new List<int>();

            if (SelectionLength == 1) return SeletResults[0];

            for (int i = 0; i < SelectionLength; i++)
            {
                if (SeletResults[i].Name == string.Empty &&
                    SeletResults[i].Sex == string.Empty &&
                    SeletResults[i].Age == string.Empty &&
                    SeletResults[i].Office == string.Empty &&
                    SeletResults[i].HosNo == string.Empty &&
                    SeletResults[i].BedNo == string.Empty &&
                    SeletResults[i].Doc == string.Empty &&
                    SeletResults[i].Retester == string.Empty)
                {
                    EmptyIndex.Add(i);
                }
                else
                {
                    NotEmptyIndex.Add(i);
                }
            }
            if (EmptyIndex.Count == SelectionLength)
            {
                return SeletResults[EmptyIndex[0]];
            }
            if (EmptyIndex.Count == SelectionLength - 1 && NotEmptyIndex.Count == 1)
            {
                return SeletResults[NotEmptyIndex[0]];
            }
            else
            {
                for (int i = 1; i < NotEmptyIndex.Count; i++)
                {
                    int CurrentRow = NotEmptyIndex[i - 1];
                    int NextRow = NotEmptyIndex[i];
                    string CurrentRowXlh = SeletResults[CurrentRow].Xlh;
                    string NextRowXlh = SeletResults[NextRow].Xlh;

                    if (SeletResults[CurrentRow].Name != SeletResults[NextRow].Name)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "姓名");
                        return null;
                    }

                    if (SeletResults[CurrentRow].Sex != SeletResults[NextRow].Sex)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "性别");
                        return null;
                    }

                    if (SeletResults[CurrentRow].Age != SeletResults[NextRow].Age)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "年龄");
                        return null;
                    }

                    if (SeletResults[CurrentRow].Office != SeletResults[NextRow].Office)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "科室");
                        return null;
                    }

                    if (SeletResults[CurrentRow].HosNo != SeletResults[NextRow].HosNo)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "门诊号");
                        return null;
                    }

                    if (SeletResults[CurrentRow].BedNo != SeletResults[NextRow].BedNo)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "床号");
                        return null;
                    }

                    if (SeletResults[CurrentRow].Doc != SeletResults[NextRow].Doc)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "医生");
                        return null;
                    }

                    if (SeletResults[CurrentRow].Retester != SeletResults[NextRow].Retester)
                    {
                        this.ShowMessageBox(CurrentRowXlh, NextRowXlh, "操作者");
                        return null;
                    }
                }
            }
            return SeletResults[NotEmptyIndex[0]];
        }

        private void ShowMessageBox(string CurrentRow, string NextRow, String ResultValueName)
        {
            MessageBox.Show("第" + CurrentRow.ToString() + "行与" + "第" + NextRow.ToString() +"行，"+ ResultValueName + "不同无法打印", 
                "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
