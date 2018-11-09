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
    
    public partial class formSign : Form
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

        private string strUserName = "";
        private string strUserDW = "";

        public formSign()
        {
            InitializeComponent();
        }

        private void formSign_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
            if (checkSoftSerial()) //已有注册
            {

            }
            else //没有注册
            {
                MessageBox.Show("未发现软件狗，请插入软件狗。", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //strVersion = "-1";
                this.Close();
                return;
            }

            
            
            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                    checkBoxPhoto.Enabled = false;
                    checkBoxSound.Checked = false;
                    checkBoxSound.Enabled = false;
                    break;
                case "2": //标准版
                    checkBoxPhoto.Enabled = false;
                    break;
                case "3": //豪华版
                    break;
                default:
                    break;
            } 
            
            labelAddi.Text = "";
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;

            boolRead = false;

            sqlDA.SelectCommand = sqlComm;
            sqlConn.Open();

            sqlComm.CommandText = "SELECT 会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间 FROM 会议表 WHERE (会议ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONFERENCE");
            labelConferenceName.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();


            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号, 参会人员表.分组名称, 参会人员表.座位号, 参会人员表.签到时间, 参会人员表.签出时间, 人员表.人员ID, 人员表.人员类别 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ")";            
            sqlDA.Fill(dSet, "PEOPLE");
            tableSign=dSet.Tables["PEOPLE"];

            sqlComm.CommandText = "SELECT 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            if (dSet.Tables.Contains("系统参数表")) dSet.Tables.Remove("系统参数表");
            sqlDA.Fill(dSet, "系统参数表");

         
            sqlConn.Close();

            InitSignView();

            intpCount = tableSign.Rows.Count;
            intpRead = dSet.Tables["SIGN"].Rows.Count;
            changeStatusBar();


        }


        private bool checkSoftSerial()
        {
            bool boolSign = true;

            ClassSenseLock cSenseLock = new ClassSenseLock();



            cSenseLock.strModal = "itconference";
            cSenseLock.iStart = 0;
            cSenseLock.iLength = 16;

            int iR = cSenseLock.ReadSenseLock();

            if (iR != 0)
                boolSign = false;
            else
                strVersion = cSenseLock.sVersion;

            return boolSign;
        }

        private void changeStatusBar()
        {
            int intpAbsent;

            intpRead = dSet.Tables["SIGN"].Rows.Count;
            intpAbsent = intpCount - intpRead;
            toolStripStatusLabelSign.Text = "会议登记：" + intpCount.ToString() + "人；已签到：" + intpRead.ToString() + "人；未签到："+intpAbsent.ToString()+"人；新增："+dataGridViewNew.Rows.Count.ToString()+"人";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            timerS.Stop();
            this.Close();
        }

        private void btnSign_Click(object sender, EventArgs e)
        {
            
            textBoxSign.Text = "";
            tabControl1.SelectedIndex = 1;
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

                switch(intCard)
                {
                    case 0:
                        textBoxSign.TextChanged+=new EventHandler(textBoxSign_TextChanged);
                        this.textBoxSign.Focus();
                        this.textBoxSign.SelectAll();
                        break;
                    case 1:
                        cMember.iComPort = intComNumber;
                        cMember.lComBand = lComBand;
                        textBoxSign.TextChanged -= new EventHandler(textBoxSign_TextChanged);
                        timerS.Start();
                        break;
                }

            }

            
            
       }

        private void textBoxSign_TextChanged(object sender, EventArgs e)
       {
            if (!boolRead) return;
            if (this.textBoxSign.Text.Length == intCardLength)
            {
               
                this.textBoxSign.SelectAll();
                ConferenceSign();

            }
        }

        private void ConferenceSign()
        {
            if (textBoxSign.Text.Trim() == "")
                return;


            //System.Data.DataRow[] signRow;
            System.Data.SqlClient.SqlDataReader sqldr;
            //签到


            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (人员表.主卡号 = '" + textBoxSign.Text + "')";
            //sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (人员表.姓名 = N'" + strUserName + "') AND (人员表.所属单位 = N'" + strUserDW + "')";
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
                textBoxZT.BackColor = Color.Red;
                textBoxZT.Text = "";
                pictureBoxPhoto.Image = null;
                sqldr.Close();
                sqlConn.Close();

                return;
            }

            sqldr.Read();
            string PID = sqldr.GetValue(14).ToString();

            
            //人员信息
            if (checkBoxPeopleInfo.Checked)
            {
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
                labelAddi.Text = "";
                int i;
                if (dSet.Tables["系统参数表"].Rows.Count > 0)
                {
                    for (i = 0; i < 4; i++)
                    {
                        if (dSet.Tables["系统参数表"].Rows[0][i].ToString() != "" && sqldr.GetValue(16 + i).ToString() != "")
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
            }
            else
                sqldr.Close();

            //是否受邀会议
            sqlComm.CommandText = "SELECT 分组名称, 座位号, 签到时间, 签出时间 FROM 参会人员表 WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ")";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)//受邀
            {
                sqldr.Read();
                labelfzmc.Text = sqldr.GetValue(0).ToString();
                labelzwh.Text = sqldr.GetValue(1).ToString();

                textBoxZT.BackColor = Color.Green;
                //textBoxZT.Text = "受邀人员";

                if (sqldr.GetValue(2).ToString() == "") //没签到
                {
                    sqldr.Close();
                    sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + System.DateTime.Now.ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ")";
                    sqlComm.ExecuteNonQuery();
                    labelwarn.Text = "受邀人员签到成功！";

                }
                else//签出
                {

                    sqldr.Close();
                    sqlComm.CommandText = "UPDATE 参会人员表 SET 签出时间 = '" + System.DateTime.Now.ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ")";
                    sqlComm.ExecuteNonQuery();
                    labelwarn.Text = "受邀人员已经签到,顺利签出！";

                }
                if (!sqldr.IsClosed) sqldr.Close();
                if (checkBoxSound.Checked)
                {
                    PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
                }
                if (intCard == 1)
                {
                    cMember.beep();
                }



            }
            else //新增
            {
                labelfzmc.Text = "";
                labelzwh.Text = "";


                textBoxZT.BackColor = Color.Orange;
                //textBoxZT.Text = "未受邀人员";
                if (!sqldr.IsClosed) sqldr.Close();
                if (checkBoxNew.Checked) //统计新增人员
                {
                    sqlComm.CommandText = "SELECT 签到时间 FROM 会议新到人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ") AND (签到时间 IS NOT NULL)";
                    sqldr = sqlComm.ExecuteReader();
                    if (!sqldr.HasRows)//没签到
                    {
                        sqldr.Close();

                        sqlComm.CommandText = "INSERT INTO 会议新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + System.DateTime.Now.ToString() + "', NULL)";
                        sqlComm.ExecuteNonQuery();
                        labelwarn.Text = "新增人员签到成功！";

                    }
                    else//签出
                    {

                        sqldr.Close();
                        sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签出时间 = '" + System.DateTime.Now.ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + PID + ")";
                        sqlComm.ExecuteNonQuery();
                        labelwarn.Text = "新增人员已经签到,顺利签出！";

                    }
                    if (!sqldr.IsClosed) sqldr.Close();
                    if (checkBoxSound.Checked)
                    {
                        PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
                    }
                    if (intCard == 1)
                    {
                        cMember.beep();
                        cMember.beep();
                    }
                }
                else//不统计新增
                {
                    labelwarn.Text = "卡无效，没有受邀参加会议！";
                }
            }


            sqlConn.Close();
            InitSignView();
            changeStatusBar();

        }

        private void textBoxSign_Leave(object sender, EventArgs e)
        {
            if (!boolRead) return;
            this.textBoxSign.Text = "";
            this.textBoxSign.Focus();
        }

        private void timerS_Tick(object sender, EventArgs e)
        {

            int iCard = 0;
            timerS.Stop();
            if (!boolRead)
            {
                timerS.Start();
                return;
            }

            iCard = cMember.getCardSerial1();
            if (iCard == 0)
            {
                textBoxSign.Text = cMember.sCardserial;

                //cMember.getUserName();
                //strUserName = cMember.strUserName;
                //strUserDW = cMember.strUserDW;

                ConferenceSign();
            }
            else
                if (iCard == -100)
                    labelwarn.Text = "该卡不是有效卡";
                textBoxSign.Text = "";

                timerS.Start();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInfo.SelectedIndices.Count; i++)
            {
                j = listBoxInfo.SelectedIndices[i];
                listBoxInput.Items.Add(listBoxInfo.Items[j]);
                listBoxInfo.Items.RemoveAt(j);

            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInput.SelectedIndices.Count; i++)
            {
                j = listBoxInput.SelectedIndices[i];
                if (listBoxInput.Items[j].ToString() == "姓名") continue;
                listBoxInfo.Items.Add(listBoxInput.Items[j]);
                listBoxInput.Items.RemoveAt(j);

            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j < 1) return;
            string strSelect = listBoxInput.Items[j - 1].ToString();
            listBoxInput.Items[j - 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j - 1]);
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j >= listBoxInput.Items.Count - 1) return;
            string strSelect = listBoxInput.Items[j + 1].ToString();
            listBoxInput.Items[j + 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j + 1]);
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            InitSignView();
        }

        private void btnDefault_Click(object sender, EventArgs e)
        {
            string []sInfo={"性别","年龄","职称","职务","所属单位","科室","手机","电话","传真","EMail","通讯地址","邮政编码","人员类别","附加属性一","附加属性二","附加属性三","附加属性四","编号"};
            string []sInput={"姓名"};
            int i;

            listBoxInfo.Items.Clear();
            listBoxInput.Items.Clear();

            for (i = 0; i < sInfo.Length; i++)
                listBoxInfo.Items.Add(sInfo[i]);

            for (i = 0; i < sInput.Length; i++)
                listBoxInput.Items.Add(sInput[i]);

        }

        private void listBoxInfo_DoubleClick(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInfo.SelectedIndices.Count; i++)
            {
                j = listBoxInfo.SelectedIndices[i];
                listBoxInput.Items.Add(listBoxInfo.Items[j]);
                listBoxInfo.Items.RemoveAt(j);

            }
        }

        private void listBoxInput_DoubleClick(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInput.SelectedIndices.Count; i++)
            {
                j = listBoxInput.SelectedIndices[i];
                if (listBoxInput.Items[j].ToString() == "姓名") continue;
                listBoxInfo.Items.Add(listBoxInput.Items[j]);
                listBoxInput.Items.RemoveAt(j);

            }
        }

        private void InitSignView()
        {
            string sTemp = "";
            int i;

            dataGridViewSign.DataSource = null;
            dataGridViewUnsign.DataSource = null;
            dataGridViewNew.DataSource = null;

            sqlConn.Open();

            sTemp = "SELECT ";
            for (i = 0; i < listBoxInput.Items.Count; i++)
            {
                sTemp += " 人员表." + listBoxInput.Items[i].ToString()+",";
            }

            sqlComm.CommandText = sTemp + " 参会人员表.签到时间, 参会人员表.签出时间 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (参会人员表.签到时间 IS NOT NULL) ORDER BY 参会人员表.签到时间";
            if (dSet.Tables.Contains("SIGN")) dSet.Tables.Remove("SIGN");
            sqlDA.Fill(dSet, "SIGN");

            sqlComm.CommandText = sTemp + " 参会人员表.签到时间 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (参会人员表.签到时间 IS NULL) ORDER BY 参会人员表.签到时间";
            if (dSet.Tables.Contains("UNSIGN")) dSet.Tables.Remove("UNSIGN");
            sqlDA.Fill(dSet, "UNSIGN");

            sqlComm.CommandText = sTemp + " 会议新到人员表.签到时间, 会议新到人员表.签出时间 FROM 人员表 INNER JOIN 会议新到人员表 ON 人员表.人员ID = 会议新到人员表.人员ID WHERE (会议新到人员表.会议ID = " + intConferenceID.ToString() + ") AND (会议新到人员表.签到时间 IS NOT NULL) ORDER BY 会议新到人员表.签到时间";
            if (dSet.Tables.Contains("NEW")) dSet.Tables.Remove("NEW");
            sqlDA.Fill(dSet, "NEW");

            sqlConn.Close();
            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dSet.Tables["SIGN"].Rows.Count>0)
                dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["SIGN"].Rows.Count - 1;
            dataGridViewUnsign.DataSource = dSet.Tables["UNSIGN"];
            dataGridViewUnsign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewNew.DataSource = dSet.Tables["NEW"];
            dataGridViewNew.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (dSet.Tables["NEW"].Rows.Count > 0)
                dataGridViewNew.FirstDisplayedScrollingRowIndex = dSet.Tables["NEW"].Rows.Count - 1;


        }

        private void formSign_FormClosed(object sender, FormClosedEventArgs e)
        {
            timerS.Stop();
        }
    }
}