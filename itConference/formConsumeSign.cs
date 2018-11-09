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
    public partial class formConsumeSign : Form
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

        
        public formConsumeSign()
        {
            InitializeComponent();
        }

        private void formConsumeSign_Load(object sender, EventArgs e)
        {
            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                case "2": //标准版
                    MessageBox.Show("该版本不支持此项功能！请购买高级版本。", "版本信息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "3": //豪华版
                    break;
                default:
                    break;
            } 
            
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;

            boolRead = false;

            sqlDA.SelectCommand = sqlComm;
            sqlConn.Open();

            sqlComm.CommandText = "SELECT 会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间 FROM 会议表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "CONFERENCE");
            labelConsume.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();

            sqlComm.CommandText = "SELECT 会议消费表.消费ID, 会议消费表.消费名称, 会议消费表.开始时间, 会议消费表.结束时间 FROM 会议表 INNER JOIN 会议消费表 ON 会议表.会议ID = 会议消费表.会议ID WHERE (会议表.会议ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONSUME");
            dataGridViewCosume.DataSource = dSet.Tables["CONSUME"];

            dataGridViewCosume.Columns[0].Visible = false;
            dataGridViewCosume.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCosume.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewCosume.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewCosume.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号, 参会人员表.分组名称, 参会人员表.座位号, 参会人员表.签到时间, 参会人员表.签出时间, 人员表.人员ID, 人员表.人员类别 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "PEOPLE");

            intpCount = dSet.Tables["PEOPLE"].Rows.Count;
            initdataGridViewSign();

            sqlConn.Close();

            if (dSet.Tables["CONSUME"].Rows.Count < 1)
            {
                MessageBox.Show("消费尚未定义！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Dispose();
                return;
            }

        }

        private void changeStatusBar()
        {
            int intpAbsent;
            intpRead = dSet.Tables["SIGN"].Rows.Count;
            intpAbsent = intpCount - intpRead;
            toolStripStatusLabelCSign.Text = "会议人数：" + intpCount.ToString() + "人；已登记：" + intpRead.ToString() + "人；未登记：" + intpAbsent.ToString() + "人";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            timerS.Stop();
        }

        private void dataGridViewCosume_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            initdataGridViewSign();
        }

        private void initdataGridViewSign()
        {
            if (dataGridViewCosume.Rows.Count < 1) return;
            dataGridViewSign.DataSource = "";

            if (dSet.Tables.Contains("SIGN")) dSet.Tables.Remove("SIGN");

            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.所属单位, 人员消费表.消费时间, 人员表.人员ID FROM 人员消费表 INNER JOIN 会议消费表 ON 人员消费表.消费ID = 会议消费表.消费ID INNER JOIN 人员表 ON 人员消费表.人员ID = 人员表.人员ID WHERE (人员消费表.消费ID = " + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ") ORDER BY 人员消费表.消费时间";
            sqlDA.Fill(dSet, "SIGN");
            tableSign = dSet.Tables["SIGN"];

            dataGridViewSign.DataSource = dSet.Tables["SIGN"];

            dataGridViewSign.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewSign.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[3].Visible = false;
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            intpRead = dSet.Tables["SIGN"].Rows.Count;
            changeStatusBar();
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            textBoxSign.Text = "";
            if (dataGridViewCosume.Rows.Count < 1) return;
            if(System.DateTime.Now<System.DateTime.Parse(dataGridViewCosume.SelectedRows[0].Cells[2].Value.ToString()))
            {
                MessageBox.Show("消费时间尚未开始！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (System.DateTime.Now > System.DateTime.Parse(dataGridViewCosume.SelectedRows[0].Cells[3].Value.ToString()))
            {
                MessageBox.Show("消费时间已结束！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (boolRead)
            {
                boolRead = false;
                btnRead.Text = "签到";
                timerS.Stop();

            }
            else
            {
                boolRead = true;
                btnRead.Text = "停止";

                switch (intCard)
                {
                    case 0:
                        textBoxSign.TextChanged += new EventHandler(textBoxSign_TextChanged);
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
                ConsumeSign();

            }
        }

        private void ConsumeSign()
        {

            //System.Data.DataRow[] signRow;
            System.Data.SqlClient.SqlDataReader sqldr;
            //签到
            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号, 参会人员表.分组名称, 参会人员表.座位号, 参会人员表.签到时间, 参会人员表.签出时间, 人员表.人员ID, 人员表.人员类别, 人员表.编号 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + " AND 主卡号 = '" + textBoxSign.Text + "')";
            sqlConn.Open();
            sqldr = sqlComm.ExecuteReader();

            if (!sqldr.HasRows)
            {
                labelwarn.Text = "卡无效，不在消费范围内";
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
                labelrylb.Text = "";
                labelBH.Text = "";
                textBoxZT.BackColor = Color.Red;
                pictureBoxPhoto.Image = null;
                sqldr.Close();
                sqlConn.Close();
                return;
            }

            textBoxZT.BackColor = Color.Green;
            //System.Data.DataRow[] sRow;
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
            labelrylb.Text = sqldr.GetValue(19).ToString();
            labelBH.Text = sqldr.GetValue(20).ToString();
            string PID = sqldr.GetValue(18).ToString();
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
                sqldr.Close();

            }
            sqlConn.Close();


            //签到检验
            sqlComm.CommandText = "SELECT 消费时间 FROM 人员消费表 WHERE (人员ID = " + PID + ") AND (消费ID = " + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ")";
            sqlConn.Open();
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Read();
                labelwarn.Text = "已经消费完毕，时间为：" + sqldr.GetValue(0).ToString();
                sqldr.Close();

                sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.所属单位, 人员消费表.消费时间, 人员表.人员ID FROM 人员消费表 INNER JOIN 会议消费表 ON 人员消费表.消费ID = 会议消费表.消费ID INNER JOIN 人员表 ON 人员消费表.人员ID = 人员表.人员ID WHERE (人员消费表.消费ID = " + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ") ORDER BY 人员消费表.消费时间";
                dSet.Tables["SIGN"].Clear(); 
                sqlDA.Fill(dSet, "SIGN");
                tableSign = dSet.Tables["SIGN"];
                sqlConn.Close();
                return;
            }
            sqldr.Close();
            sqlConn.Close();

            //labelfzmc.Text = signRow[0][14].ToString();
            //labelzwh.Text = signRow[0][15].ToString();

            sqlConn.Open();

            //string[] dgvrow ={ signRow[0][0].ToString(), signRow[0][5].ToString(), System.DateTime.Now.ToString(), signRow[0][18].ToString() };
            //dSet.Tables["SIGN"].Rows.Add(dgvrow);
            

            sqlComm.CommandText = "INSERT INTO 人员消费表 (消费ID, 人员ID, 消费时间) VALUES (" + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ", " + PID + ", '" + System.DateTime.Now.ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.所属单位, 人员消费表.消费时间, 人员表.人员ID FROM 人员消费表 INNER JOIN 会议消费表 ON 人员消费表.消费ID = 会议消费表.消费ID INNER JOIN 人员表 ON 人员消费表.人员ID = 人员表.人员ID WHERE (人员消费表.消费ID = " + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ") ORDER BY 人员消费表.消费时间";
            dSet.Tables["SIGN"].Clear();
            sqlDA.Fill(dSet, "SIGN");
            dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["SIGN"].Rows.Count - 1;

            intpRead++;
            changeStatusBar();

            labelwarn.Text = "消费完毕";


            if (checkBoxSound.Checked)
            {
                PlaySound(Directory.GetCurrentDirectory() + "\\sound.wav", 0, SND_ASYNC | SND_FILENAME);//
            }


            sqlConn.Close();
        }

        private void dataGridViewCosume_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewSign.DataSource = "";
            if (dataGridViewCosume.SelectedRows.Count < 1) return;

            if (dSet.Tables.Contains("SIGN")) dSet.Tables.Remove("SIGN");

            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.所属单位, 人员消费表.消费时间, 人员表.人员ID FROM 人员消费表 INNER JOIN 会议消费表 ON 人员消费表.消费ID = 会议消费表.消费ID INNER JOIN 人员表 ON 人员消费表.人员ID = 人员表.人员ID WHERE (人员消费表.消费ID = " + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ")";
            sqlDA.Fill(dSet, "SIGN");
            tableSign = dSet.Tables["SIGN"];

            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            intpRead = dSet.Tables["SIGN"].Rows.Count;
            changeStatusBar();
        }

        private void dataGridViewCosume_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            
            dataGridViewSign.DataSource = "";
            if (dataGridViewCosume.SelectedRows.Count < 1) return;

            if (dSet.Tables.Contains("SIGN")) dSet.Tables.Remove("SIGN");

            sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.所属单位, 人员消费表.消费时间, 人员表.人员ID FROM 人员消费表 INNER JOIN 会议消费表 ON 人员消费表.消费ID = 会议消费表.消费ID INNER JOIN 人员表 ON 人员消费表.人员ID = 人员表.人员ID WHERE (人员消费表.消费ID = " + dataGridViewCosume.SelectedRows[0].Cells[0].Value.ToString() + ")";
            sqlDA.Fill(dSet, "SIGN");
            tableSign = dSet.Tables["SIGN"];

            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            dataGridViewSign.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewSign.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[3].Visible = false;
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            intpRead = dSet.Tables["SIGN"].Rows.Count;
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
            int iCard=0;
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
                ConsumeSign();
                cMember.beep();
            }
            else
            {
                if (iCard == -100)
                    labelwarn.Text = "该卡不是有效卡";
                textBoxSign.Text = "";
            }
            timerS.Start();
        }

        private void formConsumeSign_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerS.Stop();
        }
    }
}