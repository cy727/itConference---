using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formConsumeDef : Form
    {
        public string strConn;
        public int intConferenceID;
        public string strVersion="0";
        private System.Data.DataSet dSet = new DataSet();
        
        public formConsumeDef()
        {
            InitializeComponent();
        }

        private void btnCanCel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void formConsumeDef_Load(object sender, EventArgs e)
        {

            //�汾����
            switch (strVersion)
            {
                case "1": //������
                case "2": //��׼��
                    MessageBox.Show("�ð汾��֧�ִ���ܣ��빺��߼��汾��", "�汾��Ϣ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "3": //������
                    break;
                default:
                    break;
            } 
            
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT ��������, ��ʼʱ��, ����ʱ�� FROM ����� WHERE (����ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONFERENCE");

            labelConference.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();
            dateTimePickerCstart.Value = DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][1].ToString());
            dateTimePickerCend.Value = DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][2].ToString());

            sqlComm.CommandText = "SELECT ����ID, ��������, ��ʼʱ��, ����ʱ�� FROM �������ѱ� WHERE (����ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONSUME");
            sqlConn.Close();

            dataGridViewConsume.DataSource = dSet.Tables["CONSUME"];
            dataGridViewConsume.Columns[0].Visible = false;
            dataGridViewConsume.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewConsume.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewConsume.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewConsume.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridViewConsume.Columns[0].Visible = false;


        }

        private void btncAdd_Click(object sender, EventArgs e)
        {
            if (textBoxCname.Text.Trim() == "")
            {
                MessageBox.Show("�����������Ʋ���Ϊ�գ�", "�������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.CommandText = "INSERT INTO �������ѱ� (��������, ��ʼʱ��, ����ʱ��, ����ID) VALUES (N'" + textBoxCname.Text.Trim() + "', '" + dateTimePickerCstart.Value.ToString() + "', '" + dateTimePickerCend.Value.ToString() + "', "+intConferenceID.ToString()+")";
            sqlComm.ExecuteNonQuery();

            if (dSet.Tables.Contains("CONSUME")) dSet.Tables.Remove("CONSUME");
            sqlComm.CommandText = "SELECT ����ID, ��������, ��ʼʱ��, ����ʱ�� FROM �������ѱ� WHERE (����ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "CONSUME");

            dataGridViewConsume.DataSource = dSet.Tables["CONSUME"];
            dataGridViewConsume.Columns[0].Visible = false;
            sqlConn.Close();

        }

        private void btncEdit_Click(object sender, EventArgs e)
        {
            if (textBoxCname.Text.Trim() == "")
            {
                MessageBox.Show("�����������Ʋ���Ϊ�գ�", "�������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.CommandText = "UPDATE �������ѱ� SET �������� = N'" + textBoxCname.Text.Trim() + "', ��ʼʱ�� = '" + dateTimePickerCstart.Value.ToString() + "', ����ʱ�� = '" + dateTimePickerCend.Value.ToString() + "' WHERE (����ID = " + dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString()+ ")";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();

            dataGridViewConsume.SelectedRows[0].Cells[1].Value = textBoxCname.Text.Trim();
            dataGridViewConsume.SelectedRows[0].Cells[2].Value = dateTimePickerCstart.Value;
            dataGridViewConsume.SelectedRows[0].Cells[3].Value = dateTimePickerCend.Value;

        }

        private void btncDel_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�Ƿ�ɾ������������ѣ��ù��̲��ɻ��ˡ�", "����", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.No) return;
            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM �������ѱ� WHERE (����ID = " + dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString() + ")";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();

            dataGridViewConsume.Rows.Remove(dataGridViewConsume.SelectedRows[0]);
        }

        private void textBoxCname_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCname.Text.Length > 100)
            {
                this.errorProviderC.SetError(this.textBoxCname, "�ֶι���");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderC.Clear();
            }
        }

        private void dataGridViewConsume_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewConsume.SelectedRows.Count < 1) return;
            textBoxCname.Text = dataGridViewConsume.SelectedRows[0].Cells[1].Value.ToString();
            dateTimePickerCstart.Value = DateTime.Parse(dataGridViewConsume.SelectedRows[0].Cells[2].Value.ToString());
            dateTimePickerCend.Value = DateTime.Parse(dataGridViewConsume.SelectedRows[0].Cells[3].Value.ToString());

        }
    }
}