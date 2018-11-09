using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;

namespace itConference
{
    class classCOM
    {
        public string sComPort = "COM1";
        public int lComBand = 0;
        public string sCardserial = "";

        [DllImport("libincept.dll")]
        public static extern int POS_down_file(string CommPort, string StrFile);
        [DllImport("libincept.dll")]
        public static extern int POS_up_file(string CommPort, string StrFile, string FileName);
        [DllImport("libincept.dll")]
        public static extern int POS_down_string(string CommPort, string StrString, long MaxLength);
        [DllImport("libincept.dll")]
        public static extern int POS_up_string(string CommPort, string StrString, long MaxLength);

        public int DownFile(string sInFile)
        {
            int i = 0;

            i = POS_down_file(sComPort, sInFile);//上传到POS机
            //i = POS_up_file(sComPort, s1, s2);//下载到本机
            //i = POS_down_string(sComPort, "87651234", l);
            //i = POS_up_string(sComPort, s1, l);
            return i;
        }

        public int UpFile(string sInFile)
        {
            int i = 0;

            string s1 = "", s2 = "";
            i = POS_up_file(sComPort, sInFile, s2);//下载到本机

            //i = POS_down_string(sComPort, "87651234", l);
            //i = POS_up_string(sComPort, s1, l);
            return i;
        }
    }


}
