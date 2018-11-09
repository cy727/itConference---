using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formSystemSet : Form
    {
        public string strConn;
        private System.Data.DataSet dSet = new DataSet();
        public string strCompanyName;
        public string strVersion = "0";

        private System.Windows.Forms.TextBox[] textBoxAddi;

        public formSystemSet()
        {
            InitializeComponent();

            this.textBoxAddi = new System.Windows.Forms.TextBox[4];
            this.textBoxAddi[0] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[1] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[2] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[3] = new System.Windows.Forms.TextBox();

            // 
            // textBoxAddi0
            // 
            this.textBoxAddi[0].Location = new System.Drawing.Point(15, 45);
            this.textBoxAddi[0].Name = "textBoxAddi0";
            this.textBoxAddi[0].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[0].TabIndex = 9;
            // 
            // textBoxAddi1
            // 
            this.textBoxAddi[1].Location = new System.Drawing.Point(15, 93);
            this.textBoxAddi[1].Name = "textBoxAddi1";
            this.textBoxAddi[1].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[1].TabIndex = 10;
            // 
            // textBoxAddi2
            // 
            this.textBoxAddi[2].Location = new System.Drawing.Point(15, 137);
            this.textBoxAddi[2].Name = "textBoxAddi2";
            this.textBoxAddi[2].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[2].TabIndex = 11;
            // 
            // textBoxAddi3
            // 
            this.textBoxAddi[3].Location = new System.Drawing.Point(15, 184);
            this.textBoxAddi[3].Name = "textBoxAddi3";
            this.textBoxAddi[3].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[3].TabIndex = 12;
            // 

            this.groupBox3.Controls.Add(this.textBoxAddi[3]);
            this.groupBox3.Controls.Add(this.textBoxAddi[2]);
            this.groupBox3.Controls.Add(this.textBoxAddi[1]);
            this.groupBox3.Controls.Add(this.textBoxAddi[0]);

        }

        private void formSystemSet_Load(object sender, EventArgs e)
        {
            int i;
            //�汾����
            switch (strVersion)
            {
                case "1": //������
                case "2": //��׼��
                    btnAddi.Enabled = false;
                    break;
                case "3": //������
                    break;
                default:
                    break;
            } 

            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlConn.Open();

            sqlComm.CommandText = "SELECT ID, ��˾��, ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            sqlDA.Fill(dSet, "ϵͳ������");

            if (dSet.Tables["ϵͳ������"].Rows.Count > 0)
            {
                textBoxCompanyName.Text = dSet.Tables["ϵͳ������"].Rows[0][1].ToString();
                for (i = 0; i < 4; i++)
                {
                    textBoxAddi[i].Text = dSet.Tables["ϵͳ������"].Rows[0][i + 2].ToString();
                }
            }
            else
                textBoxCompanyName.Text = "";


            sqlComm.CommandText = "SELECT ����ID, �������� FROM ��Ա�����";
            sqlDA.Fill(dSet, "��Ա�����");

            if (dSet.Tables["��Ա�����"].Rows.Count>0)
            {
                for(i=0;i<dSet.Tables["��Ա�����"].Rows.Count;i++)
                {
                    listBoxGroup.Items.Add(dSet.Tables["��Ա�����"].Rows[i][1].ToString());
                }
            }
            sqlConn.Close();
        }

        private void listBoxGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxGroup.SelectedItems.Count < 1) return;

            textBoxGroup.Text = listBoxGroup.SelectedItems[0].ToString();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if(textBoxGroup.Text.Trim()=="") return;
            listBoxGroup.Items.Add(textBoxGroup.Text.Trim());

            sqlConn.Open();
            sqlComm.CommandText = "INSERT INTO ��Ա����� (��������) VALUES (N'"+textBoxGroup.Text.Trim()+"')";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();

            MessageBox.Show("�����Ѿ���ӳɹ���");

        }

        private void btnCompanydef_Click(object sender, EventArgs e)
        {

            sqlConn.Open();
            if (dSet.Tables["ϵͳ������"].Rows.Count > 0)
            {
                sqlComm.CommandText = "UPDATE ϵͳ������ SET ��˾�� = N'" + textBoxCompanyName.Text.Trim() + "'";
            }
            else
            {
                sqlComm.CommandText = "INSERT INTO ϵͳ������ (��˾��) VALUES (N'"+textBoxCompanyName.Text.Trim()+"')";
            }
            sqlComm.ExecuteNonQuery();

            dSet.Tables["ϵͳ������"].Clear();
            sqlComm.CommandText = "SELECT ID, ��˾��, ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            sqlDA.Fill(dSet, "ϵͳ������");

            sqlConn.Close();

            MessageBox.Show("��˾�����Ѿ��޸�Ϊ����" + textBoxCompanyName.Text.Trim()+"��");
            strCompanyName = textBoxCompanyName.Text.Trim();

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if(textBoxGroup.Text.Trim()=="") return;
            if (listBoxGroup.SelectedItems.Count < 1)
            {
                MessageBox.Show("��ѡ��Ҫ�޸ĵķ��飡", "����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBox.Show("�Ƿ��޸ķ������ƣ��ù��̲��ɻ��ˣ�", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            
            sqlConn.Open();

            sqlComm.CommandText = "UPDATE ��Ա����� SET �������� = N'" + textBoxGroup.Text + "' WHERE (�������� = N'" + listBoxGroup.SelectedItems[0].ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlComm.CommandText = "UPDATE ��Ա�� SET ��Ա��� = N'" + textBoxGroup.Text + "' WHERE (��Ա��� = N'" + listBoxGroup.SelectedItems[0].ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
            listBoxGroup.Items[listBoxGroup.SelectedIndices[0]] = textBoxGroup.Text;

            MessageBox.Show("���������Ѿ��޸ģ�");

        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (listBoxGroup.SelectedItems.Count < 1)
            {
                MessageBox.Show("��ѡ��Ҫɾ���ķ��飡", "����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBox.Show("�Ƿ�ɾ�����飿�ù��̲��ɻ��ˣ�", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;


            sqlConn.Open();

            sqlComm.CommandText = "DELETE FROM ��Ա����� WHERE (�������� = N'" + listBoxGroup.SelectedItems[0] .ToString()+ "')";
            sqlComm.ExecuteNonQuery();

            sqlComm.CommandText = "UPDATE ��Ա�� SET ��Ա��� = NULL WHERE (��Ա��� = N'" + listBoxGroup.SelectedItems[0].ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
            listBoxGroup.Items.Remove(listBoxGroup.SelectedItems[0]);
            listBoxGroup.SelectedItems.Clear();

            MessageBox.Show("���������Ѿ�ɾ����");
        }

        private void textBoxCompanyName_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCompanyName.Text.Length > 100)
            {
                this.errorProviders.SetError(this.textBoxCompanyName, "�ֶι���");
                e.Cancel = true;
            }
            else
            {
                this.errorProviders.Clear();
            }
        }

        private void textBoxGroup_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxGroup.Text.Length > 100)
            {
                this.errorProviders.SetError(this.textBoxGroup, "�ֶι���");
                e.Cancel = true;
            }
            else
            {
                this.errorProviders.Clear();
            }
        }

        private void btnAddi_Click(object sender, EventArgs e)
        {

            sqlConn.Open();
            if (dSet.Tables["ϵͳ������"].Rows.Count > 0)
            {
                sqlComm.CommandText = "UPDATE ϵͳ������ SET ��������һ = N'" + textBoxAddi[0].Text.Trim() + "', �������Զ� = N'" + textBoxAddi[1].Text.Trim() + "', ���������� = N'" + textBoxAddi[2].Text.Trim() + "', ���������� = N'"+textBoxAddi[3].Text.Trim()+"'";
            }
            else
            {
                sqlComm.CommandText = "INSERT INTO ϵͳ������ (��������һ, �������Զ�, ����������, ����������) VALUES (N'" + textBoxAddi[0].Text.Trim() + "', N'" + textBoxAddi[1].Text.Trim() + "', N'" + textBoxAddi[2].Text.Trim() + "', N'" + textBoxAddi[3].Text.Trim() + "')";
            }
            sqlComm.ExecuteNonQuery();

            dSet.Tables["ϵͳ������"].Clear();
            sqlComm.CommandText = "SELECT ID, ��˾��, ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            sqlDA.Fill(dSet, "ϵͳ������");

            sqlConn.Close();

            MessageBox.Show("��Ա���������Ѿ�����");

        }
    }
}