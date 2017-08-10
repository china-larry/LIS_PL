using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace FineCareLisSYS
{

   public class BaseFormAuto :Form
    {
        #region 控件缩放
        double formWidth;//窗体原始宽度
        double formHeight;//窗体原始高度
        double scaleX;//水平缩放比例
        double scaleY;//垂直缩放比例

        public BaseFormAuto()
        {

        }

        Dictionary<string, string> controlInfo = new Dictionary<string, string>();//控件中心Left,Top,控件Width,控件Height,控件字体Size
        /// <summary>
        /// 获取所有原始数据
        /// </summary>
        protected void GetAllInitInfo(Control CrlContainer)
        {
            if (CrlContainer.Parent == null)//CrlContainer是主窗体
            {
                formWidth = Convert.ToDouble(CrlContainer.Width);
                formHeight = Convert.ToDouble(CrlContainer.Height);
            }
            foreach (Control item in CrlContainer.Controls)
            {
                if (item.Name.Trim() != "")
                    controlInfo.Add(item.Name, (item.Left + item.Width / 2) + "," + (item.Top + item.Height / 2) + "," + item.Width + "," + item.Height + "," + item.Font.Size);
                if ((item as UserControl) == null && item.Controls.Count > 0)//是控件中含有控件
                    GetAllInitInfo(item);//递归调用
            }
        }
        public void ControlsChangeInit(Control CrlContainer)
        {
            scaleX = (Convert.ToDouble(CrlContainer.Width) / formWidth);
            scaleY = (Convert.ToDouble(CrlContainer.Height) / formHeight);
        }
        public void ControlsChange(Control CrlContainer)
        {
            double[] pos = new double[5];//pos数组保存当前控件中心X,Y,控件Width,控件Height,控件字体Size
            foreach (Control item in CrlContainer.Controls)
            {
                if (item.Name.Trim() != "")
                {
                    if ((item as UserControl) == null && item.Controls.Count > 0)//用户控件
                        ControlsChange(item);
                    string[] strs = controlInfo[item.Name].Split(',');//获取键值
                    for (int j = 0; j < 5; j++)
                    {
                        pos[j] = Convert.ToDouble(strs[j]);
                    }
                    double itemWidth = pos[2] * scaleX;//宽度放大
                    double itemHeight = pos[3] * scaleY;
                    item.Left = Convert.ToInt32(pos[0] * scaleX - itemWidth / 2);//计算放大后的左
                    item.Top = Convert.ToInt32(pos[1] * scaleY - itemHeight / 2);
                    item.Width = Convert.ToInt32(itemWidth);
                    item.Height = Convert.ToInt32(itemHeight);
                    if (item.Name.IndexOf("StringGrid")==-1)//StringGrid表格字体不变
                    {
                        item.Font = new Font(item.Font.Name, float.Parse((pos[4] * Math.Min(scaleX, scaleY)).ToString()));
                    }
                    
                }
            }
        }
        #endregion 


       

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (controlInfo.Count > 0)
            {
                ControlsChangeInit(this);//获取缩放比例.Controls[0]
                ControlsChange(this);//缩放.Controls[0]
            }
        }
    
    }


}
