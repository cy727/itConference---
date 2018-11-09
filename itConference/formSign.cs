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
        public static extern bool PlaySound(string pszSound, int hmod, int fdwSound);//����windows���֣�����
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
            if (checkSoftSerial()) //����ע��
            {

            }
            else //û��ע��
            {
                MessageBox.Show("δ�����������������������", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //strVersion = "-1";
                this.Close();
                return;
            }

            
            
            //�汾����
            switch (strVersion)
            {
                case "1": //������
                    checkBoxPhoto.Enabled = false;
                    checkBoxSound.Checked = false;
                    checkBoxSound.Enabled = false;
                    break;
                case "2": //��׼��
                    checkBoxPhoto.Enabled = false;
                    break;
                case "3": //������
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

            sqlComm.CommandText = "SELECT ��������, ���鸱��, ��������, �᳡��, ��ʼʱ��, ����ʱ�� FROM ����� WHERE (����ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONFERENCE");
            labelConferenceName.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();


            sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������, �λ���Ա��.��������, �λ���Ա��.��λ��, �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ��, ��Ա��.��ԱID, ��Ա��.��Ա��� FROM ��Ա�� INNER JOIN �λ���Ա�� ON ��Ա��.��ԱID = �λ���Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ")";            
            sqlDA.Fill(dSet, "PEOPLE");
            tableSign=dSet.Tables["PEOPLE"];

            sqlComm.CommandText = "SELECT ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            if (dSet.Tables.Contains("ϵͳ������")) dSet.Tables.Remove("ϵͳ������");
            sqlDA.Fill(dSet, "ϵͳ������");

         
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
            toolStripStatusLabelSign.Text = "����Ǽǣ�" + intpCount.ToString() + "�ˣ���ǩ����" + intpRead.ToString() + "�ˣ�δǩ����"+intpAbsent.ToString()+"�ˣ�������"+dataGridViewNew.Rows.Count.ToString()+"��";
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
                    if (MessageBox.Show("������δ��ʼ�Ƿ�ʼǩ����", "ȷ��ǩ��", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                        return;
                }
                if (System.DateTime.Now > System.DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][5].ToString()))
                {
                    if (MessageBox.Show("�����Ѿ��������Ƿ����ǩ����", "ȷ��ǩ��", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                        return;
                }
            }
            
            
            if (boolRead)
            {
                boolRead = false;
                btnSign.Text = "ǩ��";
                timerS.Stop();

            }
            else
            {
                boolRead = true;
                btnSign.Text = "ֹͣ";

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
            //ǩ��


            sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������,��Ա��.��ԱID, ��Ա��.��Ա���, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.����������, ��Ա��.��� FROM ��Ա�� WHERE (��Ա��.������ = '" + textBoxSign.Text + "')";
            //sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������,��Ա��.��ԱID, ��Ա��.��Ա���, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.����������, ��Ա��.��� FROM ��Ա�� WHERE (��Ա��.���� = N'" + strUserName + "') AND (��Ա��.������λ = N'" + strUserDW + "')";
            sqlConn.Open();
            sqldr = sqlComm.ExecuteReader();
            //signRow = tableSign.Select("������ = '" + textBoxSign.Text + "'");

            if (!sqldr.HasRows) //�޿���
            {
                labelwarn.Text = "����Ч��û����Ա��¼��";
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

            
            //��Ա��Ϣ
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
                if (dSet.Tables["ϵͳ������"].Rows.Count > 0)
                {
                    for (i = 0; i < 4; i++)
                    {
                        if (dSet.Tables["ϵͳ������"].Rows[0][i].ToString() != "" && sqldr.GetValue(16 + i).ToString() != "")
                        {
                            labelAddi.Text = labelAddi.Text + dSet.Tables["ϵͳ������"].Rows[0][i].ToString() + ":" + sqldr.GetValue(16 + i).ToString() + "��";
                        }
                    }
                }
                labelBH.Text = sqldr.GetValue(20).ToString();

                sqldr.Close();

                if (checkBoxPhoto.Checked)
                {

                    sqlComm.CommandText = "SELECT ��Ƭ FROM ��Ƭ�� WHERE (��ԱID = " + PID + ")";
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

            //�Ƿ���������
            sqlComm.CommandText = "SELECT ��������, ��λ��, ǩ��ʱ��, ǩ��ʱ�� FROM �λ���Ա�� WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��ԱID = " + PID + ")";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)//����
            {
                sqldr.Read();
                labelfzmc.Text = sqldr.GetValue(0).ToString();
                labelzwh.Text = sqldr.GetValue(1).ToString();

                textBoxZT.BackColor = Color.Green;
                //textBoxZT.Text = "������Ա";

                if (sqldr.GetValue(2).ToString() == "") //ûǩ��
                {
                    sqldr.Close();
                    sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + System.DateTime.Now.ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID = " + PID + ")";
                    sqlComm.ExecuteNonQuery();
                    labelwarn.Text = "������Աǩ���ɹ���";

                }
                else//ǩ��
                {

                    sqldr.Close();
                    sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + System.DateTime.Now.ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID = " + PID + ")";
                    sqlComm.ExecuteNonQuery();
                    labelwarn.Text = "������Ա�Ѿ�ǩ��,˳��ǩ����";

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
            else //����
            {
                labelfzmc.Text = "";
                labelzwh.Text = "";


                textBoxZT.BackColor = Color.Orange;
                //textBoxZT.Text = "δ������Ա";
                if (!sqldr.IsClosed) sqldr.Close();
                if (checkBoxNew.Checked) //ͳ��������Ա
                {
                    sqlComm.CommandText = "SELECT ǩ��ʱ�� FROM �����µ���Ա�� WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID = " + PID + ") AND (ǩ��ʱ�� IS NOT NULL)";
                    sqldr = sqlComm.ExecuteReader();
                    if (!sqldr.HasRows)//ûǩ��
                    {
                        sqldr.Close();

                        sqlComm.CommandText = "INSERT INTO �����µ���Ա�� (����ID, ��ԱID, ǩ��ʱ��, ǩ��ʱ��) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + System.DateTime.Now.ToString() + "', NULL)";
                        sqlComm.ExecuteNonQuery();
                        labelwarn.Text = "������Աǩ���ɹ���";

                    }
                    else//ǩ��
                    {

                        sqldr.Close();
                        sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + System.DateTime.Now.ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID = " + PID + ")";
                        sqlComm.ExecuteNonQuery();
                        labelwarn.Text = "������Ա�Ѿ�ǩ��,˳��ǩ����";

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
                else//��ͳ������
                {
                    labelwarn.Text = "����Ч��û�������μӻ��飡";
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
                    labelwarn.Text = "�ÿ�������Ч��";
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
                if (listBoxInput.Items[j].ToString() == "����") continue;
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
            string []sInfo={"�Ա�","����","ְ��","ְ��","������λ","����","�ֻ�","�绰","����","EMail","ͨѶ��ַ","��������","��Ա���","��������һ","�������Զ�","����������","����������","���"};
            string []sInput={"����"};
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
                if (listBoxInput.Items[j].ToString() == "����") continue;
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
                sTemp += " ��Ա��." + listBoxInput.Items[i].ToString()+",";
            }

            sqlComm.CommandText = sTemp + " �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM ��Ա�� INNER JOIN �λ���Ա�� ON ��Ա��.��ԱID = �λ���Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (�λ���Ա��.ǩ��ʱ�� IS NOT NULL) ORDER BY �λ���Ա��.ǩ��ʱ��";
            if (dSet.Tables.Contains("SIGN")) dSet.Tables.Remove("SIGN");
            sqlDA.Fill(dSet, "SIGN");

            sqlComm.CommandText = sTemp + " �λ���Ա��.ǩ��ʱ�� FROM ��Ա�� INNER JOIN �λ���Ա�� ON ��Ա��.��ԱID = �λ���Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (�λ���Ա��.ǩ��ʱ�� IS NULL) ORDER BY �λ���Ա��.ǩ��ʱ��";
            if (dSet.Tables.Contains("UNSIGN")) dSet.Tables.Remove("UNSIGN");
            sqlDA.Fill(dSet, "UNSIGN");

            sqlComm.CommandText = sTemp + " �����µ���Ա��.ǩ��ʱ��, �����µ���Ա��.ǩ��ʱ�� FROM ��Ա�� INNER JOIN �����µ���Ա�� ON ��Ա��.��ԱID = �����µ���Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (�����µ���Ա��.ǩ��ʱ�� IS NOT NULL) ORDER BY �����µ���Ա��.ǩ��ʱ��";
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