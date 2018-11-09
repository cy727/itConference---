using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;

namespace itConference
{
    public partial class formTrainSign : Form
    {
        public string strConn;
        public int intConferenceID;
        private bool boolRead;
        private System.Data.DataSet dSet = new DataSet();
        private System.Data.DataTable tableSign;
        private int intpCount, intpRead;
        public string strVersion = "0";
        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;
        private ClassMember cMember = new ClassMember();

        [DllImport("winmm.dll")]
        public static extern bool PlaySound(string pszSound, int hmod, int fdwSound);//播放windows音乐，重载
        public const int SND_FILENAME = 0x00020000;
        public const int SND_ASYNC = 0x0001;

        private string strUserName="";
        private string strUserDW = "";

        public formTrainSign()
        {
            InitializeComponent();
        }

        private void formTrainSign_Load(object sender, EventArgs e)
        {

            int i,istart,iend;
            string strTemp;

            labelAddi.Text = "";
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;

            boolRead = false;

            sqlDA.SelectCommand = sqlComm;
            sqlConn.Open();

            sqlComm.CommandText = "SELECT 会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间 FROM 培训表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "CONFERENCE");
            labelConferenceName.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();


            sqlComm.CommandText = "SELECT 培训人员表.参会单位, 培训人员表.参会人数, 培训人员表.到会人数, 单位表.ID, 单位表.单位编号, 单位表.上级单位 FROM 培训人员表 LEFT OUTER JOIN 单位表 ON 培训人员表.参会单位 = 单位表.单位名称 WHERE (培训人员表.会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "PEOPLE");
            tableSign = dSet.Tables["PEOPLE"];

            sqlComm.CommandText = "SELECT 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            if (dSet.Tables.Contains("系统参数表")) dSet.Tables.Remove("系统参数表");
            sqlDA.Fill(dSet, "系统参数表");


            sqlComm.CommandText = "SELECT ID, 单位名称, 上级单位, 单位编号 FROM 单位表 ORDER BY 上级单位";
            if (dSet.Tables.Contains("单位表")) dSet.Tables.Remove("单位表");
            sqlDA.Fill(dSet, "单位表");

            //单位表整理
            for (i = 0; i < dSet.Tables["单位表"].Rows.Count; i++)
            {
                strTemp=dSet.Tables["单位表"].Rows[i][2].ToString();
                istart=strTemp.IndexOf(',');

                if(istart==-1) //主目录
                    dSet.Tables["单位表"].Rows[i][3]=dSet.Tables["单位表"].Rows[i][0].ToString();
                else
                {
                    istart++;
                    iend=strTemp.IndexOf(',',istart);
                    if(iend==-1)
                        iend=strTemp.Length;
;

                    dSet.Tables["单位表"].Rows[i][3]=strTemp.Substring(istart,iend-istart);
                }

            }

            //培训整理
            for (i = 0; i < dSet.Tables["PEOPLE"].Rows.Count; i++)
            {
                if (dSet.Tables["PEOPLE"].Rows[i][3].ToString() == "")
                    continue;

                DataRow[] dr = dSet.Tables["单位表"].Select("ID=" + dSet.Tables["PEOPLE"].Rows[i][3].ToString());
                if (dr.Length > 0)
                {
                    dSet.Tables["PEOPLE"].Rows[i][4] = dr[0][3].ToString();
                }

            }

            sqlConn.Close();

            InitSignView();

            intpCount = 0;
            for (i = 0; i < dSet.Tables["PEOPLE"].Rows.Count; i++)
                intpCount += (int)dSet.Tables["PEOPLE"].Rows[i][1];

            intpRead = dSet.Tables["SIGN"].Rows.Count;
            changeStatusBar();

        }

        private void changeStatusBar()
        {
            int intpAbsent;

            intpRead = dSet.Tables["SIGN"].Rows.Count;
            intpAbsent = intpCount - intpRead;
            toolStripStatusLabelSign.Text = "会议登记：" + intpCount.ToString() + "人；已签到：" + intpRead.ToString() + "人；未签到：" + intpAbsent.ToString() + "人；新增：" + dataGridViewNew.Rows.Count.ToString() + "人";
        }

        private void InitSignView()
        {
            string sTemp = "";
            int i;

            dataGridViewSign.DataSource = null;
            dataGridViewUnsign.DataSource = null;
            dataGridViewNew.DataSource = null;

            sqlConn.Open();

            sqlComm.CommandText = "SELECT 培训签到表.人员ID, 人员表.姓名, 培训签到表.代表单位, 培训签到表.签到时间, 培训签到表.签出时间 FROM 培训签到表 INNER JOIN 人员表 ON 培训签到表.人员ID = 人员表.人员ID WHERE (培训签到表.会议ID = " + intConferenceID.ToString() + ")  ORDER BY 培训签到表.签到时间";
            if (dSet.Tables.Contains("SIGN")) dSet.Tables.Remove("SIGN");
            sqlDA.Fill(dSet, "SIGN");

            sqlComm.CommandText = "SELECT 参会单位, 参会人数 - 到会人数 AS 缺席人数 FROM 培训人员表 WHERE (会议ID = "+intConferenceID.ToString()+") AND (参会人数 > 到会人数)";
            if (dSet.Tables.Contains("UNSIGN")) dSet.Tables.Remove("UNSIGN");
            sqlDA.Fill(dSet, "UNSIGN");

            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.所属单位, 培训新到人员表.签到时间, 培训新到人员表.签出时间 FROM 培训新到人员表 INNER JOIN 人员表 ON 培训新到人员表.人员ID = 人员表.人员ID WHERE (培训新到人员表.会议ID = "+intConferenceID.ToString()+") ORDER BY 培训新到人员表.签到时间";
            if (dSet.Tables.Contains("NEW")) dSet.Tables.Remove("NEW");
            sqlDA.Fill(dSet, "NEW");

            sqlConn.Close();
            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dSet.Tables["SIGN"].Rows.Count > 0)
                dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["SIGN"].Rows.Count - 1;
            dataGridViewUnsign.DataSource = dSet.Tables["UNSIGN"];
            dataGridViewUnsign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewNew.DataSource = dSet.Tables["NEW"];
            dataGridViewNew.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dSet.Tables["NEW"].Rows.Count > 0)
                dataGridViewNew.FirstDisplayedScrollingRowIndex = dSet.Tables["NEW"].Rows.Count - 1;


        }

        private void timerS_Tick(object sender, EventArgs e)
        {
            int iCard = 0;

            if (!boolRead) return;

            iCard = cMember.getCardSerial1();
            cMember.beep();
            if (iCard == 0)
            {
                textBoxSign.Text = cMember.sCardserial;
                cMember.getUserName();
                strUserName = cMember.strUserName;
                strUserDW = cMember.strUserDW;

                ConferenceSign();
            }
            else
                if (iCard == -100)
                    labelwarn.Text = "该卡不是有效卡";
            textBoxSign.Text = "";
            

            if (strUserName != "")
            {
                ConferenceSign();
            }
            else
                labelwarn.Text = "该卡不是有效卡";
        }

        private void btnSign_Click(object sender, EventArgs e)
        {
            textBoxSign.Text = "";
            tabControl1.SelectedIndex = 0;
            tabControl1.SelectedIndex = 0;
            if (!boolRead)
            {
                if (System.DateTime.Now < System.DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][4].ToString()))
                {
                    if (MessageBox.Show("会议尚未开始是否开始签到？", "确认签到", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                        return;
                }
                if (System.DateTime.Now > System.DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][5].ToString()))
                {
                    if (MessageBox.Show("会议已经结束，是否继续签到？", "确认签到", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                        return;
                }
            }


            if (boolRead)
            {
                boolRead = false;
                btnSign.Text = "签到";
                timerS.Stop();

            }
            else
            {
                boolRead = true;
                btnSign.Text = "停止";

                switch (intCard)
                {
                    case 0:
                        break;
                    case 1:
                        cMember.iComPort = intComNumber;
                        cMember.lComBand = lComBand;
                        //textBoxSign.TextChanged -= new EventHandler(textBoxSign_TextChanged);
                        timerS.Start();
                        break;
                }

            }
        }

        private void ConferenceSign()
        {
            string strTemp = "";

            //System.Data.DataRow[] signRow;
            System.Data.SqlClient.SqlDataReader sqldr;
            //签到


            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (人员表.姓名 = N'" + strUserName + "') AND (人员表.所属单位 = N'" + strUserDW +"')";
            sqlConn.Open();
            sqldr = sqlComm.ExecuteReader();
            //signRow = tableSign.Select("主卡号 = '" + textBoxSign.Text + "'");

            if (!sqldr.HasRows) //无卡号
            {
                labelwarn.Text = "卡无效，没有人员纪录！";
                labelxm.Text = "";
                labelxb.Text = "";
                labelnl.Text = "";
                labelzc.Text = "";
                labelzw.Text = "";
                labelssdw.Text = "";
                labelks.Text = "";
                labelsj.Text = "";
                labeldh.Text = "";
                labelcz.Text = "";
                labelem.Text = "";
                labeltxdz.Text = "";
                labelyzbm.Text = "";
                labelfzmc.Text = "";
                labelzwh.Text = "";
                labelryfl.Text = "";
                labelAddi.Text = "";
                pictureBoxPhoto.Image = null;
                sqldr.Close();
                sqlConn.Close();

                return;
            }

            //人员信息
            sqldr.Read();
            labelxm.Text = sqldr.GetValue(0).ToString();
            labelxb.Text = sqldr.GetValue(1).ToString();
            labelnl.Text = sqldr.GetValue(2).ToString();
            labelzc.Text = sqldr.GetValue(3).ToString();
            labelzw.Text = sqldr.GetValue(4).ToString();
            labelssdw.Text = sqldr.GetValue(5).ToString();
            labelks.Text = sqldr.GetValue(6).ToString();
            labelsj.Text = sqldr.GetValue(7).ToString();
            labeldh.Text = sqldr.GetValue(8).ToString();
            labelcz.Text = sqldr.GetValue(9).ToString();
            labelem.Text = sqldr.GetValue(10).ToString();
            labeltxdz.Text = sqldr.GetValue(11).ToString();
            labelyzbm.Text = sqldr.GetValue(12).ToString();
            labelryfl.Text = sqldr.GetValue(15).ToString();
            string PID = sqldr.GetValue(14).ToString();
            labelAddi.Text = "";
            int i;
            if (dSet.Tables["系统参数表"].Rows.Count > 0)
            {
                for (i = 0; i < 4; i++)
                {
                    if (dSet.Tables["系统参数表"].Rows[0][i].ToString() != "" && sqldr.GetValue(20 + i).ToString() != "")
                    {
                        labelAddi.Text = labelAddi.Text + dSet.Tables["系统参数表"].Rows[0][i].ToString() + ":" + sqldr.GetValue(16 + i).ToString() + "　";
                    }
                }
            }
            labelBH.Text = sqldr.GetValue(20).ToString();

            sqldr.Close();

            if (checkBoxPhoto.Checked)
            {

                sqlComm.CommandText = "SELECT 照片 FROM 照片表 WHERE (人员ID = " + PID + ")";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    try
                    {
                        byte[] bytePhoto = (byte[])sqldr.GetValue(0);
                        MemoryStream StreamPhoto = new MemoryStream(bytePhoto);
                        this.pictureBoxPhoto.Image = Image.FromStream(StreamPhoto);

                    }
                    catch
                    {
                    }
                }
                else
                {
                    pictureBoxPhoto.Image = null;
                }
                sqldr.Close();

            }
            //签出
            sqlComm.CommandText = "SELECT ID FROM 培训签到表 WHERE (人员ID = "+PID+") AND (会议ID = "+intConferenceID.ToString()+")";
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Close();
                sqlComm.CommandText = "UPDATE 参会人员表 SET 签出时间 = '" + System.DateTime.Now.ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ")";
                sqlComm.ExecuteNonQuery();
                labelwarn.Text = "受邀人员已经签到,顺利签出！";
                if (checkBoxSound.Checked)
                {
                    PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
                }


                sqlConn.Close();
                InitSignView();
                changeStatusBar();
                return;
            }
            else
                sqldr.Close();

            sqlComm.CommandText = "SELECT ID FROM 培训新到人员表 WHERE (人员ID = " + PID + ") AND (会议ID = " + intConferenceID.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Close();
                sqlComm.CommandText = "UPDATE 培训新到人员表 SET 签出时间 = '" + System.DateTime.Now.ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ")";
                sqlComm.ExecuteNonQuery();
                labelwarn.Text = "新增人员已经签到,顺利签出！";
                if (checkBoxSound.Checked)
                {
                    PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
                }


                sqlConn.Close();
                InitSignView();
                changeStatusBar();
                return;
            }
            else
                sqldr.Close();

            //本单位签到
            sqlComm.CommandText = "SELECT ID FROM 培训人员表 WHERE (会议ID = "+intConferenceID.ToString()+") AND (参会单位 = N'"+strUserDW+"') AND (参会人数 > 到会人数)";
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Close();
                sqlComm.CommandText = "INSERT INTO 培训签到表 (会议ID, 人员ID, 代表单位, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", N'" + strUserDW + "', '" + System.DateTime.Now.ToString() + "', NULL)";
                sqlComm.ExecuteNonQuery();


                sqlComm.CommandText = "UPDATE 培训人员表 SET 到会人数 = 到会人数 + 1 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + strUserDW + "')";
                sqlComm.ExecuteNonQuery();


                labelwarn.Text = "受邀人员签到成功！";
                if (checkBoxSound.Checked)
                {
                    PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
                }


                sqlConn.Close();
                InitSignView();
                changeStatusBar();
                return;

            }
            else
                sqldr.Close();

            //代表单位签到
            while (true)
            {
                if (strUserDW == "")
                    break;


                DataRow[] dr = dSet.Tables["单位表"].Select("单位名称='" + strUserDW+"'");
                if (dr.Length < 1)
                    break;
                
                strTemp=dr[0][3].ToString();

                DataRow[] dr1 = dSet.Tables["PEOPLE"].Select("单位编号= '" + strTemp +"'");
                if (dr1.Length < 1)
                    break;

                strTemp=dr1[0][0].ToString();//代表单位

                sqlComm.CommandText = "SELECT ID FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + strTemp + "') AND (参会人数 > 到会人数)";
                sqldr = sqlComm.ExecuteReader();
                if (sqldr.HasRows) //有代表单位
                {
                    sqldr.Close();
                    sqlComm.CommandText = "INSERT INTO 培训签到表 (会议ID, 人员ID, 代表单位, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", N'" + strTemp + "', '" + System.DateTime.Now.ToString() + "', NULL)";
                    sqlComm.ExecuteNonQuery();


                    sqlComm.CommandText = "UPDATE 培训人员表 SET 到会人数 = 到会人数 + 1 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + strTemp + "')";
                    sqlComm.ExecuteNonQuery();


                    labelwarn.Text = "受邀人员签到成功！";
                    if (checkBoxSound.Checked)
                    {
                        PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
                    }


                    sqlConn.Close();
                    InitSignView();
                    changeStatusBar();
                    return;

                }
                else //没有代表单位
                {
                    sqldr.Close();
                    break;
                }


            }
            
            //新增
            if (!sqldr.IsClosed) sqldr.Close();
            sqlComm.CommandText = "INSERT INTO 培训新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + System.DateTime.Now.ToString() + "', NULL)";
            sqlComm.ExecuteNonQuery();
            labelwarn.Text = "新增人员签到成功！";
            if (checkBoxSound.Checked)
            {
                PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
            }

            sqlConn.Close();
            InitSignView();
            changeStatusBar();

        }
    }
}
