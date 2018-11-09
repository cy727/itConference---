using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace itConference
{
    class ClassMember
    {
        public int iComPort = 0;
        public int lComBand = 0;
        public string sCardserial = "";

        public string strUserName = "";
        public string strUserDW = "";
        public string strUserBH = "";
        public string strUserBM = "";
        public string strUserZW = "";

        private const string sSky = "FFFFFFFFFFFF";
        //private const int BLOCKLENGTH = 5;
        //private const int BYTESPERBLOCK = 16;

        private const int BLOCKLENGTH = 8;
        private const int BYTESPERBLOCK = 16;
        private const int BYTESPERPASS = 7;

        //private byte[] bSky = { 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x0 };
        //private byte[] bSky = { 0x13, 0x69, 0x10, 0x08, 0x66, 0x2f, 0x0 };
        private byte[] bSky = { 0x13, 0x69, 0x10, 0x08, 0x66, 0x2f };
        private byte[] bSky1 = { 0xff, 0xff, 0xff, 0xff, 0xff, 0xff };
        private byte[] ibKeyBlock = { 11, 15, 19, 39, 43, 47, 51, 55, 59, 63 };
        //private byte[] ibKeyBlock = { 7 };

        private byte ibNameBlock = 8;
        private byte ibTeleBlock = 9;
        private byte ibStyleBlock = 12;
        private byte ibMoneyBlock = 13;

        private byte[] iDWBlock = {44,45,46,48,49,50};
        private byte[] iBMBlock = {52,53,54};
        private byte[] iZWBlock = {56,57,58};
        

        //private byte[] ibMMBlock = { 20, 24, 28, 32, 40, 60 };
        //private byte[] ibLBlock = { 16 };

        private byte[] ibMMBlock = { };
        private byte[] ibLBlock = { };
        private byte[] ibLBlock1 = { 16 };

        private string strLPass = "IT922";
        private string strMMPass = "IT";


        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_init_com(int port, int baud);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_request(int icdev, byte model, byte[] pTagType);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_anticoll(int icdev, byte bcnt, byte[] pSnr, byte[] pLen);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_ClosePort();
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_beep(int icdev, int msec);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_select(int icdev, byte[] pSnr, byte snrLen, byte[] pSize);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_authentication2(int icdev, byte model, byte block, byte[] pKey);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_read(int icdev, byte block, byte[] pData, byte[] pLen);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_write(int icdev, byte block, byte[] pData);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_initval(int icdev, byte block, long value);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_readval(int icdev, byte block, byte[] pValue);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_increment(int icdev, byte block, long value);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_decrement(int icdev, byte block, long value);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_restore(int icdev, byte block);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_M1_transfer(int icdev, byte block);
        [DllImport("MasterRD.dll", CharSet = CharSet.Auto)]
        public static extern int rf_halt(int icdev);

        //

        public string GetCommandString(string strInput)
        {
            if (strInput == "" || strInput == null)
            {
                return ("NULL");
            }
            else
                return strInput;
        }

        //检查，初始化,成功0
        public int checkCard()
        {
            return checkCard_S50();
        }

        public int checkCard1()
        {
            return checkCard1_S50();
        }

        public void getUserName()
        {
            strUserName = "";
            strUserDW = "";
            strUserBH = "";
            strUserBM = "";
            strUserZW = "";

            int iStatus, i, j;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];
            byte[] byteTemp = new byte[200];
            byte[] byteREAD = new byte[20];
            byte[] iTType = new byte[4];

            byte[] bTempDW = new byte[BYTESPERBLOCK * iDWBlock.Length];
            byte[] bTempBM = new byte[BYTESPERBLOCK * iBMBlock.Length];
            byte[] bTempZW = new byte[BYTESPERBLOCK * iZWBlock.Length];

            bool bVCard = true;

            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return ;

            //寻卡
            iStatus = rf_request(0x0, 0x52, iTType);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return ;
            }

            iStatus = rf_anticoll(0, 4, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return ;
            }
            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return ;
            }

            //得到用户名称
            System.Text.UTF8Encoding u8Encoding = new UTF8Encoding();
            iStatus = rf_M1_authentication2(0, 0x60, ibNameBlock, bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return ;
            }

            iStatus = rf_M1_read(0x0, ibNameBlock, byteREAD, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return ;
            }

            
            strUserName = System.Text.Encoding.Default.GetString(byteREAD);
            strUserName = strUserName.Replace("\0","");

            //编号
            /*
             iStatus = rf_M1_read(0x0, ibTeleBlock, byteREAD, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return;
            }
            strUserBH = System.Text.Encoding.Default.GetString(byteREAD);
            strUserBH = strUserBH.Replace("\0", "");
             */

            //单位
            iStatus = rf_M1_authentication2(0, 0x60, iDWBlock[0], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return ;
            }
            strUserDW = "";
            for (i = 0; i < 3; i++)
            {
                iStatus = rf_M1_read(0x0, iDWBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return;
                }
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    bTempDW[i * BYTESPERBLOCK + j] = byteREAD[j];
                }
                
            }

            iStatus = rf_M1_authentication2(0, 0x60, iDWBlock[3], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return;
            }
            for (i = 3; i < 6; i++)
            {
                iStatus = rf_M1_read(0x0, iDWBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return;
                }
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    bTempDW[i * BYTESPERBLOCK + j] = byteREAD[j];
                }
            }
            strUserDW += System.Text.Encoding.Default.GetString(bTempDW);
            strUserDW = strUserDW.Replace("\0", "");

            //部门
            /*
            iStatus = rf_M1_authentication2(0, 0x60, iBMBlock[0], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return;
            }
            strUserBM = "";
            for (i = 0; i < 3; i++)
            {
                iStatus = rf_M1_read(0x0, iBMBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return;
                }
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    bTempBM[i * BYTESPERBLOCK + j] = byteREAD[j];
                }

            }
            strUserBM += System.Text.Encoding.Default.GetString(bTempBM);
            strUserBM = strUserBM.Replace("\0", "");
             */

            //职务
            /*
            iStatus = rf_M1_authentication2(0, 0x60, iZWBlock[0], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return;
            }
            strUserZW = "";
            for (i = 0; i < 3; i++)
            {
                iStatus = rf_M1_read(0x0, iZWBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return;
                }
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    bTempZW[i * BYTESPERBLOCK + j] = byteREAD[j];
                }
            }
            strUserZW += System.Text.Encoding.Default.GetString(bTempZW);
            strUserZW = strUserZW.Replace("\0", "");
             */
           
            rf_ClosePort();
            return ;


        }

        //S50卡检查 -1
        private int checkCard_S50()
        {
            int iStatus, i;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];
            byte[] byteTemp = new byte[200];
            byte[] byteREAD = new byte[20];
            byte[] iTType = new byte[4];

            bool bVCard = true;

            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return -1;


            //寻卡
            iStatus = rf_request(0x0, 0x52, iTType);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            iStatus = rf_anticoll(0, 4, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            //检查加密区

            System.Text.UTF8Encoding u8Encoding = new UTF8Encoding();
            for (i = 0; i < ibLBlock.Length; i++)
            {
                iStatus = rf_M1_authentication2(0, 0x60, ibLBlock[i], bSky);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                iStatus = rf_M1_read(0x0, ibLBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                string strTemp = System.Text.Encoding.Default.GetString(byteREAD);
                if (strTemp.Substring(0, strLPass.Length) != strLPass) //无效卡
                {
                    bVCard = false;
                    break;
                }

            }

            if (!bVCard)
            {
                rf_ClosePort();
                MessageBox.Show("此卡为无效卡！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            for (i = 0; i < ibMMBlock.Length; i++)
            {
                iStatus = rf_M1_authentication2(0, 0x60, ibMMBlock[i], bSky);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                iStatus = rf_M1_read(0x0, ibMMBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                string strTemp1 = System.Text.Encoding.Default.GetString(byteREAD);
                if (strTemp1.Substring(0, strMMPass.Length) != strMMPass) //无效卡
                {
                    bVCard = false;
                    break;
                }

            }

            if (!bVCard)
            {
                rf_ClosePort();
                MessageBox.Show("此卡为无效卡！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            rf_ClosePort();
            return 0;
        }

        //S50卡检查 -1
        private int checkCard1_S50()
        {
            int iStatus, i;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];
            byte[] byteTemp = new byte[200];
            byte[] byteREAD = new byte[20];
            byte[] iTType = new byte[4];

            bool bVCard = true;

            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return -1;


            //寻卡
            iStatus = rf_request(0x0, 0x52, iTType);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            iStatus = rf_anticoll(0, 4, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            //检查加密区

            System.Text.UTF8Encoding u8Encoding = new UTF8Encoding();
            for (i = 0; i < ibLBlock1.Length; i++)
            {
                iStatus = rf_M1_authentication2(0, 0x60, ibLBlock1[i], bSky);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return -100;
                }

                iStatus = rf_M1_read(0x0, ibLBlock1[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                //string strTemp = System.Text.Encoding.Default.GetString(byteREAD);
                //if (strTemp.Substring(0, strLPass.Length) != strLPass) //无效卡
                //{
                //    bVCard = false;
                //    break;
                //}

            }

            if (!bVCard)
            {
                rf_ClosePort();
                return -100;
            }

            for (i = 0; i < ibMMBlock.Length; i++)
            {
                iStatus = rf_M1_authentication2(0, 0x60, ibMMBlock[i], bSky);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                iStatus = rf_M1_read(0x0, ibMMBlock[i], byteREAD, sLen);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

                string strTemp1 = System.Text.Encoding.Default.GetString(byteREAD);
                if (strTemp1.Substring(0, strMMPass.Length) != strMMPass) //无效卡
                {
                    bVCard = false;
                    break;
                }

            }

            if (!bVCard)
            {
                rf_ClosePort();
                return -100;
            }

            rf_ClosePort();
            return 0;
        }


        //读卡序列号,寻卡,成功0
        public int getCardSerial1()
        {
            sCardserial = "";
            if (iComPort == 0 || lComBand == 0) return -1;
            int iTemp = checkCard1();
            if (iTemp == -1 || iTemp==-100) return iTemp;
            return _gerCardSerial_S50();
        }

        //读卡序列号,寻卡,成功0
        public int getCardSerial()
        {
            sCardserial = "";
            if (iComPort == 0 || lComBand == 0) return -1;
            if (checkCard() == -1) return -1;
            return _gerCardSerial_S50();
        }

        //充值
        public int chargeCard(float fSum)
        {
            return chargeCard_S50(fSum);
        }

        //开卡，初始化,成功0(姓名,手机号，卡类型，余额)
        public int initCard(string sCName, string sCNumber, int iCardStyle, float fSum, string sDW, string sBM, string sZW)
        {
            //if (sCName == "" || sCNumber == "")
            if (sCName == "")
                return -1;

            if (checkCard() == -1) return -1;

            return initCard_S50(sCName, sCNumber, iCardStyle, fSum, sDW, sBM, sZW);
        }

        //S50卡充值
        private int chargeCard_S50(float fSum)
        {
            int iStatus, i;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];
            byte[] byteTemp = new byte[200];

            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return -1;

            //充值
            iStatus = rf_anticoll(0, 4, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            iStatus = rf_M1_authentication2(0, 0x60, ibStyleBlock, bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            byte[] bTemp3 = System.Text.Encoding.Default.GetBytes(fSum.ToString());
            iStatus = rf_M1_write(0x0, ibMoneyBlock, bTemp3);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            rf_ClosePort();
            return 0;


        }


        //S50卡初始化
        private int initCard_S50(string sCName, string sCNumber, int iCardStyle, float fSum, string sDW, string sBM, string sZW)
        {
            int iStatus, i,j;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];
            byte[] byteTemp = new byte[200];
            string sTemp = "",sTemp1="";
            bool bLoop;
            byte[] bTempT = new byte[BYTESPERBLOCK];



            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return -1;

            //写入会员信息
            /*
            iStatus = rf_anticoll(0, 4, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
             */
            iStatus = rf_M1_authentication2(0, 0x60, ibNameBlock, bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }


            //写入姓名
            System.Text.UTF8Encoding u8Encoding = new UTF8Encoding();
            if(sCName.Length>BLOCKLENGTH)
                sCName=sCName.Substring(BLOCKLENGTH);

            byte[] bTemp =new byte[BYTESPERBLOCK];

            bTemp=System.Text.Encoding.Default.GetBytes(sCName);

            for (i = 0; i < BYTESPERBLOCK; i++)
            {
                if (i >= bTemp.Length)
                    bTempT[i] = 0x0;
                else
                    bTempT[i] = bTemp[i];
            }
            iStatus = rf_M1_write(0x0, ibNameBlock, bTempT);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            byte[] bTemp1 = System.Text.Encoding.Default.GetBytes(sCNumber);
            //iStatus = rf_M1_read(0x0, 0x05, btemp2,btemp3);
            iStatus = rf_M1_write(0x0, ibTeleBlock, bTemp1);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            //写入卡类型
            iStatus = rf_M1_authentication2(0, 0x60, ibStyleBlock, bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            byte[] bTemp2 = System.Text.Encoding.Default.GetBytes(iCardStyle.ToString());
            //iStatus = rf_M1_read(0x0, 0x05, btemp2,btemp3);
            iStatus = rf_M1_write(0x0, ibStyleBlock, bTemp2);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            //初始化钱包
            byte[] bTemp3 = System.Text.Encoding.Default.GetBytes(fSum.ToString());
            iStatus = rf_M1_write(0x0, ibMoneyBlock, bTemp3);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            //写入单位
            byte[] bTemp5 = new byte[BYTESPERBLOCK * iDWBlock.Length];
            if (sDW.Length > BLOCKLENGTH * iDWBlock.Length)
                sTemp = sDW.Substring(BLOCKLENGTH * iDWBlock.Length);
            else
                sTemp = sDW;

            bTemp5 = System.Text.Encoding.Default.GetBytes(sTemp);

            iStatus = rf_M1_authentication2(0, 0x60, iDWBlock[0], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            for (i = 0; i < 3; i++)
            {
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    if(i * BYTESPERBLOCK + j>=bTemp5.Length)
                        bTempT[j]=0x0;
                    else
                        bTempT[j] = bTemp5[i * BYTESPERBLOCK + j];
                }
                //写入
                iStatus = rf_M1_write(0x0, iDWBlock[i], bTempT);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

            }

            iStatus = rf_M1_authentication2(0, 0x60, iDWBlock[3], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            for (i = 3; i < 6; i++)
            {
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    if(i * BYTESPERBLOCK + j>=bTemp5.Length)
                        bTempT[j]=0x0;
                    else
                        bTempT[j] = bTemp5[i * BYTESPERBLOCK + j];
                }
                //写入
                iStatus = rf_M1_write(0x0, iDWBlock[i], bTempT);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

            }



            //写入部门
            /*
            byte[] bTemp6 = new byte[BYTESPERBLOCK * iBMBlock.Length];

            if (sBM.Length > BLOCKLENGTH * iBMBlock.Length)
                sTemp = sBM.Substring(BLOCKLENGTH * iBMBlock.Length);
            else
                sTemp = sBM;

            bTemp6 = System.Text.Encoding.Default.GetBytes(sTemp);

            iStatus = rf_M1_authentication2(0, 0x60, iBMBlock[0], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            for (i = 0; i < 3; i++)
            {
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    if (i * BYTESPERBLOCK + j >= bTemp6.Length)
                        bTempT[j] = 0x0;
                    else
                        bTempT[j] = bTemp6[i * BYTESPERBLOCK + j];
                }
                //写入
                iStatus = rf_M1_write(0x0, iBMBlock[i], bTempT);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

            }
            */


            //写入职务
            /*
            byte[] bTemp4 = new byte[BYTESPERBLOCK * iZWBlock.Length];
 
            if (sZW.Length > BLOCKLENGTH * iZWBlock.Length)
                sTemp = sZW.Substring(BLOCKLENGTH * iZWBlock.Length);
            else
                sTemp = sZW;
            bTemp4 = System.Text.Encoding.Default.GetBytes(sTemp);

            iStatus = rf_M1_authentication2(0, 0x60, iZWBlock[0], bSky);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            for (i = 0; i < 3; i++)
            {
                for (j = 0; j < BYTESPERBLOCK; j++)
                {
                    if (i * BYTESPERBLOCK + j >= bTemp4.Length)
                        bTempT[j] = 0x0;
                    else
                        bTempT[j] = bTemp4[i * BYTESPERBLOCK + j];
                }
                //写入
                iStatus = rf_M1_write(0x0, iZWBlock[i], bTempT);
                if (iStatus != 0)
                {
                    rf_ClosePort();
                    return iStatus;
                }

            }
            */

            rf_ClosePort();
            return 0;
        }


        //S50卡读卡序列号
        private int _gerCardSerial_S50()
        {
            int iStatus, i;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];

            byte[] iTType = new byte[4];
            Encoding enc = Encoding.Default;
            byte[] byteTemp = new byte[200];

            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return -1;

            //寻卡
            iStatus = rf_request(0x0, 0x52, iTType);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            //BEEP
            /*
            iStatus = rf_beep(0x0,10);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            */

            //序列号
            iStatus = rf_anticoll(0x0, 0x04, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            sCardserial = "";
            string sCardtemp = "";

            for (i = 0; i < Int32.Parse(sLen[0].ToString()); i++)
            {
                sCardtemp = Convert.ToString(sSerial[i], 16).ToUpper();
                if (sCardtemp.Length < 2)
                    sCardtemp = "0" + sCardtemp;
                //sCardserial += Convert.ToString(sSerial[i], 16).ToUpper();
                sCardserial += sCardtemp;
            }

            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            rf_ClosePort();
            return 0;
        }


        private long alterNumber(long lInput)
        {
            string sNumber = "";
            //转换为16进制
            string sTemp = lInput.ToString("X");

            if (sTemp.Length > 8)
                return 0;

            //补齐字节
            sTemp = sTemp.PadLeft(8, '0');

            //字节转顺序
            int i;
            for (i = 1; i <= 4; i++)
            {
                sNumber = sNumber + sTemp.Substring(sTemp.Length - i * 2, 2);
            }

            //
            return long.Parse(sNumber, System.Globalization.NumberStyles.AllowHexSpecifier);
        }

        public void beep()
        {
            beep_S50();
        }

        private void beep_S50()
        {
            int iStatus, i;

            //打开端口
            rf_init_com(iComPort, lComBand);
            rf_beep(0x0, 10);
            rf_ClosePort();
        }

        public int addKey()
        {
            int iStatus, i, j;
            byte[] sLen = new byte[1];
            byte[] sSerial = new byte[200];
            byte[] byteTemp = new byte[200];
            byte[] bTempT = new byte[BYTESPERBLOCK];
            byte[] iTType = new byte[4];


            //打开端口
            iStatus = rf_init_com(iComPort, lComBand);
            if (iStatus != 0) return iStatus;

            //寻卡
            iStatus = rf_request(0x0, 0x52, iTType);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            iStatus = rf_anticoll(0, 4, sSerial, sLen);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }
            iStatus = rf_select(0, sSerial, sLen[0], byteTemp);
            if (iStatus != 0)
            {
                rf_ClosePort();
                return iStatus;
            }

            bool bVail = true;
            for (i = 0; i < ibKeyBlock.Length; i++)
            {
                iStatus = rf_M1_authentication2(0, 0x60, ibKeyBlock[i], bSky1);
                if (iStatus != 0)
                {
                    //rf_ClosePort();
                    bVail = false;
                    continue;
                }


                //加密
                iStatus = rf_M1_read(0x0, ibKeyBlock[i], bTempT, sLen);
                if (iStatus != 0)
                {
                    //rf_ClosePort();
                    bVail = false;
                    continue;
                }
                for (j = 0; j < bSky.Length; j++)
                    bTempT[j] = bSky[j];

                iStatus = rf_M1_write(0x0, ibKeyBlock[i], bTempT);
                if (iStatus != 0)
                {
                    //rf_ClosePort();
                    bVail = false;
                    continue;
                }
            }

            rf_ClosePort();
            if (bVail)
                return 0;
            else
                return -100;
        }

        public string IDCARDNOFROMHEX(string idNo, int noLong)
        {
            string st="";
            int itemp=0;

            if (idNo == "")
                return st;

            itemp = int.Parse(idNo, System.Globalization.NumberStyles.AllowHexSpecifier);

            return itemp.ToString("D"+noLong.ToString());
        }


    }
}
