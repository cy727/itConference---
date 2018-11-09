using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using SmsATCommdll;

namespace itConference
{
    public partial class formTZ : Form
    {
        private System.Data.DataSet dSet1 = new DataSet();
        SmsATComm at = new SmsATComm();
        formConference fc;

        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        private int iMAXSTRING = 70;
        private int iMAXSMS = 3;


        public formTZ()
        {
            InitializeComponent();
        }

        public formTZ(formConference parentForm)
        {
            InitializeComponent();
            this.fc = parentForm;

            sqlConn.ConnectionString = fc.strConn;
            sqlComm.Connection = sqlConn;
        }

        private void formTZ_Load(object sender, EventArgs e)
        {
            string dFileName1 = Directory.GetCurrentDirectory() + "\\addcon.xml";

            if (File.Exists(dFileName1)) //存在文件
            {
                dSet1.ReadXml(dFileName1);
                comboBoxCOM.SelectedIndex = Int32.Parse(dSet1.Tables["附加信息"].Rows[0][1].ToString());
            }
            else
            {
                comboBoxCOM.SelectedIndex = 0;
            }

            //fc = (formConference)this.ParentForm;
            textBoxTZ.Text = "请您于" + fc.dateTimePickerStart.Value.ToLongDateString() + " " + fc.dateTimePickerStart.Value.ToShortTimeString() + " 参加 " + fc.textBoxCname.Text.Trim();
            if (fc.textBoxCname2.Text.Trim() != "")
            {
                textBoxTZ.Text += "(" + fc.textBoxCname2.Text.Trim() + ")";
            }
            if (fc.textBoxConferenceNo.Text.Trim() != "")
            {
                textBoxTZ.Text += " 会场：" + fc.textBoxConferenceNo.Text.Trim();
            }

            textBoxTZ.Text += " ，请准时到会。";

            toolStripStatusLabelCount.Text = "0/"+fc.intCount.ToString();


        }

        private void btnDOWN_Click(object sender, EventArgs e)
        {

            int iCount,i,iSMS=0,j ;
            string[] strTZ = {"","",""};


            if (textBoxTZ.Text.Trim().Length == 0)
            {
                MessageBox.Show("请输入短信内容！", "会议通知", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (textBoxTZ.Text.Trim().Length > iMAXSMS * iMAXSTRING)
            {
                MessageBox.Show("短信过长，请适当缩短短信内容！", "会议通知", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            for(i=0;i<iMAXSMS;i++)
            {
                if (textBoxTZ.Text.Trim().Substring(i*iMAXSTRING).Length <= iMAXSTRING)
                {
                    strTZ[i] = textBoxTZ.Text.Trim().Substring(i * iMAXSTRING);
                    break;
                }

                strTZ[i] = textBoxTZ.Text.Trim().Substring(i * iMAXSTRING,iMAXSTRING);
            
            }

            iSMS=i+1;

            if (radioButtonWTZ.Checked) //未通知
                iCount = fc.intWTZ;
            else
                iCount = fc.intCount;

            if (iCount == 0)
            {
                MessageBox.Show("没有需要发送的信息！", "会议通知", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            toolStripStatusLabelCount.Text = "0/" + iCount.ToString();
            toolStripProgressBarCount.Maximum=iCount;
            toolStripProgressBarCount.Value = 0;
            toolStripStatusLabelStatus.Text = "";

            
            i = 0;

            at.portNames = comboBoxCOM.Text;
            at.BaudRate = 9600;
            at.OpenPort();
            //at.DataReceivedChanged += new SmsATComm.SMS_DataReceivedEventHandler(at_DataReceivedChanged);

            sqlConn.Open();

            bool bSUCCESS;

            foreach (DataGridViewRow dgvP in fc.dataGridViewPeople.Rows)
            {
                if (!bool.Parse(dgvP.Cells[0].Value.ToString())) //不参加
                    continue;

                if (dgvP.Cells[9].Value.ToString().Trim() == "") //手机为空
                {
                    textBoxRZ.Text += "发送失败：" + dgvP.Cells[2].Value.ToString() + "（手机空号）\r\n";
                    i++;
                    if (i > iCount)
                        i = iCount;
                    toolStripStatusLabelCount.Text = i.ToString()+"/" + iCount.ToString();
                    toolStripProgressBarCount.Value = i;
                    toolStripStatusLabelStatus.Text = "";
                    continue;
                }

                if(radioButtonWTZ.Checked) //未通知
                    if (dgvP.Cells[11].Value.ToString() != "") //已通知
                    {
                        continue;
                    }


                toolStripStatusLabelStatus.Text = "正在发送:" + dgvP.Cells[2].Value.ToString() + "（" + dgvP.Cells[9].Value.ToString() + "）";
                this.Refresh();

                bSUCCESS = true;
                for (j = 0; j < iSMS; j++)
                {
                    if (!at.SendSMS(dgvP.Cells[9].Value.ToString(), strTZ[j]))
                    {
                        bSUCCESS = false;
                        break;
                    }
                }


                if (bSUCCESS) //成功短信
                {
                    sqlComm.CommandText = "UPDATE 参会人员表 SET 短信 = '" + System.DateTime.Now.ToString() + "' WHERE (人员ID = " + dgvP.Cells[1].Value.ToString() + ") AND (会议ID = " + fc.intConferenceID.ToString() + ") ";
                    sqlComm.ExecuteNonQuery();
                    textBoxRZ.Text += "发送成功：" + dgvP.Cells[2].Value.ToString() + "（" + dgvP.Cells[9].Value.ToString() + "）\r\n";
                }
                else //不成功短信
                {
                    sqlComm.CommandText = "UPDATE 参会人员表 SET 短信 = NULL  WHERE (人员ID = " + dgvP.Cells[1].Value.ToString() + ") AND (会议ID = " + fc.intConferenceID.ToString() + ") ";
                    sqlComm.ExecuteNonQuery();
                    textBoxRZ.Text += "发送失败：" + dgvP.Cells[2].Value.ToString() + "（" + dgvP.Cells[9].Value.ToString() + "）\r\n";
                }
                i++;
                if (i > iCount)
                    i = iCount;
                toolStripStatusLabelCount.Text = i.ToString() + "/" + iCount.ToString();
                toolStripProgressBarCount.Value = i;

            }
            sqlConn.Close();
            at.ClosePort();

            MessageBox.Show("通知发送完毕！", "会议通知", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            toolStripStatusLabelStatus.Text = "发送完毕";
            
        }

        private void btnRZ_Click(object sender, EventArgs e)
        {

            saveFileDialogIO.Filter = "TXT files(*.txt)|*.txt";
            saveFileDialogIO.FileName = "通知记录" + System.DateTime.Now.ToShortDateString();

            if (saveFileDialogIO.ShowDialog() != DialogResult.OK) return;
            this.Refresh();

            FileStream mStream = new FileStream(saveFileDialogIO.FileName, FileMode.Create);
            StreamWriter mWriter = new StreamWriter(mStream, Encoding.UTF8);
            mWriter.Write(textBoxRZ.Text);
            mWriter.Flush();
            mWriter.Close();
            mStream.Close();
        }
    }
}
