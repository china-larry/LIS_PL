using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections;
using System.Windows.Forms;

namespace FineCareLisSYS
{
    /// <summary>
    /// IniFiles的类
    /// </summary>
    public class RWIni
    {
        public static string FileName = Application.StartupPath + "\\Config.ini"; //INI文件名
        //声明读写INI文件的API函数
        [DllImport("kernel32")]
        private static extern bool WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, byte[] retVal, int size, string filePath);
        //类的构造函数，传递INI文件名
        public RWIni()
        {
            string AFileName = FileName;
            // 判断文件是否存在
            FileInfo fileInfo = new FileInfo(AFileName);
            //Todo:搞清枚举的用法
            if (!fileInfo.Exists)   //如果ini文件不存在,则新建一个?
            { //文件不存在，建立文件
                //MessageBox.Show("ini配置文件不存在,请检查文件路径!","提示");
                StreamWriter sw = new StreamWriter(AFileName, false, Encoding.Default);
                try
                {
                    //sw.Write("#参数配置");
                    sw.Close();
                }

                catch
                {
                    throw (new ApplicationException("Ini文件不存在"));
                }
            }
            //必须是完全路径，不能是相对路径
            //FileName = fileInfo.FullName;
        }
        //写INI文件
        public static void WriteString(string Section, string Ident, string Value)
        {
            if (!File.Exists(FileName))   //如果ini文件不存在,则新建一个?
            { //文件不存在，建立文件
                FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs);
                sw.WriteLine("[" + Section + "]");
                sw.WriteLine(Ident + "=");
                sw.Close();
                fs.Close();
            }
            if (!WritePrivateProfileString(Section, Ident, Value, FileName))
            {
                throw (new ApplicationException("写Ini文件出错"));
            }
        }
        //读取INI文件指定
        public static string ReadString(string Section, string Ident, string Default)
        {     
            Byte[] Buffer = new Byte[65535];
            int bufLen = GetPrivateProfileString(Section, Ident, Default, Buffer, Buffer.GetUpperBound(0), FileName);
            //必须设定0（系统默认的代码页）的编码方式，否则无法支持中文
            string s = Encoding.GetEncoding(0).GetString(Buffer);
            s = s.Substring(0, bufLen);
            return s.Trim();
        }

        //读整数
        public static int ReadInteger(string Section, string Ident, int Default)
        {
            string intStr = ReadString(Section, Ident, Convert.ToString(Default));
            try
            {
                return Convert.ToInt32(intStr);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Default;
            }
        }
        //写整数
        public static void WriteInteger(string Section, string Ident, int Value)
        {
            WriteString(Section, Ident, Value.ToString());
        }



        //读布尔
        public static bool ReadBool(string Section, string Ident, bool Default)
        {
            try
            {
                return Convert.ToBoolean(ReadString(Section, Ident, Convert.ToString(Default)));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Default;
            }
        }

        //写Bool
        public static void WriteBool(string Section, string Ident, bool Value)
        {
            WriteString(Section, Ident, Convert.ToString(Value));
        }

    }
}

