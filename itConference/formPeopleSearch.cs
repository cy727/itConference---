using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formPeopleSearch : Form
    {
        public string strConn;
        public int intReturn;
        public System.Windows.Forms.TextBox[] textBoxAddi;
        private System.Windows.Forms.Label[] labelAddi;
        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;
        private ClassMember cMember = new ClassMember();

        public formPeopleSearch()
        {
            InitializeComponent();

            this.labelAddi = new System.Windows.Forms.Label[4];
            this.labelAddi[0] = new System.Windows.Forms.Label();
            this.labelAddi[1] = new System.Windows.Forms.Label();
            this.labelAddi[2] = new System.Windows.Forms.Label();
            this.labelAddi[3] = new System.Windows.Forms.Label();
            this.textBoxAddi = new System.Windows.Forms.TextBox[4];
            this.textBoxAddi[0] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[1] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[2] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[3] = new System.Windows.Forms.TextBox();

            // 
            // labelAddi0
            // 
            this.labelAddi[0].AutoSize = true;
            this.labelAddi[0].Location = new System.Drawing.Point(8, 21);
            this.labelAddi[0].Name = "labelAddi0";
            this.labelAddi[0].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[0].TabIndex = 0;
            this.labelAddi[0].Text = "��������һ��";
            // 
            // labelAddi2
            // 
            this.labelAddi[2].AutoSize = true;
            this.labelAddi[2].Location = new System.Drawing.Point(8, 50);
            this.labelAddi[2].Name = "labelAddi2";
            this.labelAddi[2].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[2].TabIndex = 1;
            this.labelAddi[2].Text = "������������";
            // 
            // labelAddi1
            // 
            this.labelAddi[1].AutoSize = true;
            this.labelAddi[1].Location = new System.Drawing.Point(195, 21);
            this.labelAddi[1].Name = "labelAddi1";
            this.labelAddi[1].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[1].TabIndex = 2;
            this.labelAddi[1].Text = "�������Զ���";
            // 
            // labelAddi3
            // 
            this.labelAddi[3].AutoSize = true;
            this.labelAddi[3].Location = new System.Drawing.Point(195, 51);
            this.labelAddi[3].Name = "labelAddi3";
            this.labelAddi[3].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[3].TabIndex = 3;
            this.labelAddi[3].Text = "���������ģ�";
            // 
            // textBoxAddi0
            // 
            this.textBoxAddi[0].Location = new System.Drawing.Point(82, 18);
            this.textBoxAddi[0].Name = "textBoxAddi0";
            this.textBoxAddi[0].Size = new System.Drawing.Size(100, 21);
            this.textBoxAddi[0].TabIndex = 4;
            // 
            // textBoxAddi1
            // 
            this.textBoxAddi[1].Location = new System.Drawing.Point(268, 18);
            this.textBoxAddi[1].Name = "textBoxAddi1";
            this.textBoxAddi[1].Size = new System.Drawing.Size(100, 21);
            this.textBoxAddi[1].TabIndex = 5;
            // 
            // textBoxAddi2
            // 
            this.textBoxAddi[2].Location = new System.Drawing.Point(82, 46);
            this.textBoxAddi[2].Name = "textBoxAddi2";
            this.textBoxAddi[2].Size = new System.Drawing.Size(100, 21);
            this.textBoxAddi[2].TabIndex = 6;
            // 
            // textBoxAddi3
            // 
            this.textBoxAddi[3].Location = new System.Drawing.Point(268, 46);
            this.textBoxAddi[3].Name = "textBoxAddi3";
            this.textBoxAddi[3].Size = new System.Drawing.Size(100, 21);
            this.textBoxAddi[3].TabIndex = 7;

            this.groupBox3.Controls.Add(this.textBoxAddi[3]);
            this.groupBox3.Controls.Add(this.textBoxAddi[2]);
            this.groupBox3.Controls.Add(this.textBoxAddi[1]);
            this.groupBox3.Controls.Add(this.textBoxAddi[0]);
            this.groupBox3.Controls.Add(this.labelAddi[3]);
            this.groupBox3.Controls.Add(this.labelAddi[1]);
            this.groupBox3.Controls.Add(this.labelAddi[2]);
            this.groupBox3.Controls.Add(this.labelAddi[0]);

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            intReturn = 0;
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (textBoxUsername.Text.Trim() == "" && comboBoxGender.Text.Trim() == "" && textBoxPPost.Text.Trim() == "" && textBoxPost.Text.Trim() == "" && textBoxCompany.Text.Trim() == "" && textBoxAddress.Text.Trim() == "" && comboBoxGroup.Text.Trim() == "" && textBoxCard.Text.Trim() == "" && textBoxAddi[0].Text.Trim() == "" && textBoxAddi[1].Text.Trim() == "" && textBoxAddi[2].Text.Trim() == "" && textBoxAddi[3].Text.Trim() == "")
                intReturn = -1;
            else
                intReturn = 1;

            this.Close();
        }

        private void formPeopleSearch_Load(object sender, EventArgs e)
        {
            int i;
            this.CancelButton = this.btnCancel;


            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;
            //������Ա��Ϣ
            sqlConn.Open();
            sqlComm.CommandText = "SELECT ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            if (dSet.Tables.Contains("ϵͳ������")) dSet.Tables.Remove("ϵͳ������");
            sqlDA.Fill(dSet, "ϵͳ������");
            sqlConn.Close();

            if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
            {
                for (i = 0; i < 4; i++)
                {
                    labelAddi[i].Text = "��������:";
                    textBoxAddi[i].Text = "";
                    textBoxAddi[i].Enabled = false;
                }
            }
            else
            {
                for (i = 0; i < 4; i++)
                {
                    if (dSet.Tables["ϵͳ������"].Rows[0][i].ToString() == "")
                    {
                        labelAddi[i].Text = "��������:";
                        textBoxAddi[i].Text = "";
                        textBoxAddi[i].Enabled = false;
                    }
                    else
                    {
                        labelAddi[i].Text = dSet.Tables["ϵͳ������"].Rows[0][i].ToString() + ":";
                        textBoxAddi[i].Text = "";
                        textBoxAddi[i].Enabled = true;
                    }
                }

            }

            sqlComm.CommandText = "SELECT �������� FROM ��Ա�����";
            if (dSet.Tables.Contains("��Ա�����")) dSet.Tables.Remove("��Ա�����");
            sqlDA.Fill(dSet, "��Ա�����");
            sqlConn.Close();

            for (i = 0; i < dSet.Tables["��Ա�����"].Rows.Count; i++)
              {
                  comboBoxGroup.Items.Add(dSet.Tables["��Ա�����"].Rows[i][0]);
              }
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            switch (intCard)
            {
                case 0:
                    this.textBoxCard.Focus();
                    this.textBoxCard.SelectAll();
                    break;
                case 1:
                    cMember.iComPort = intComNumber;
                    cMember.lComBand = lComBand;
                    if (cMember.getCardSerial() == 0)
                        textBoxCard.Text = cMember.sCardserial;
                    else
                        textBoxCard.Text = "";
                    cMember.beep();
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            intReturn = -1;
            this.Close();
        }
        
    }
}