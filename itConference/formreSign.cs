using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formreSign : Form
    {
        public string strConn;
        public int intConferenceID;
        private System.Data.DataSet dSet = new DataSet();

        public formreSign()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void formreSign_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;


            sqlDA.SelectCommand = sqlComm;
            sqlConn.Open();

            sqlComm.CommandText = "SELECT ��������, ��ʼʱ��, ����ʱ�� FROM ����� WHERE (����ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONFERENCE");
            //dataGridViewreSign.DataSource = dSet.Tables["CONFERENCE"];

            sqlComm.CommandText = "SELECT ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            if (dSet.Tables.Contains("ϵͳ������")) dSet.Tables.Remove("ϵͳ������");
            sqlDA.Fill(dSet, "ϵͳ������");


            sqlComm.CommandText = "SELECT ��Ա��.��ԱID, ��Ա��.����, ��Ա��.������λ, ��Ա��.ְ��, ��Ա��.ְ��, �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ��, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.���������� FROM ��Ա�� INNER JOIN �λ���Ա�� ON ��Ա��.��ԱID = �λ���Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ")  AND (�λ���Ա��.ǩ��ʱ�� IS NULL)";
            sqlDA.Fill(dSet, "ReSIGN");
            dataGridViewreSign.DataSource = dSet.Tables["ReSIGN"];
            dataGridViewreSign.Columns[0].Visible = false;

            dataGridViewreSign.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewreSign.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewreSign.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewreSign.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewreSign.Columns[5].Visible = false;
            dataGridViewreSign.Columns[6].Visible = false;
            dataGridViewreSign.Columns[7].Visible = false;
            dataGridViewreSign.Columns[8].Visible = false;
            dataGridViewreSign.Columns[9].Visible = false;
            dataGridViewreSign.Columns[10].Visible = false;
            dataGridViewreSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            labelConferencename.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();

            changeStatusBar();
            sqlConn.Close();
        }

        private void changeStatusBar()
        {

            labelAddi.Text = "";

            toolStripStatusLabelresign.Text = "��ѡ����" + dataGridViewreSign .SelectedRows.Count+ "��";

            if (dataGridViewreSign.SelectedRows.Count<1)
            {
                return;
            }
            
            int i;
            if (dSet.Tables["ϵͳ������"].Rows.Count > 0)
            {
                for (i = 0; i < 4; i++)
                {
                    if (dSet.Tables["ϵͳ������"].Rows[0][i].ToString() != "" && dataGridViewreSign.SelectedRows[0].Cells[7+i].Value.ToString()!= "")
                    {
                        labelAddi.Text = labelAddi.Text + dSet.Tables["ϵͳ������"].Rows[0][i].ToString() + ":" + dataGridViewreSign.SelectedRows[0].Cells[7 + i].Value.ToString() + "��";
                    }
                }
            }

        }

        private void btnreSign_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�Ƿ�Ҫ���в�ǩ���ù��̲��ɻ���!", "��Ϣ��ʾ", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                return;

            sqlConn.Open();
            int i;
            for(i=0;i<dataGridViewreSign.SelectedRows.Count;i++)
            {
                sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '"+dSet.Tables["CONFERENCE"].Rows[0][1].ToString()+"', ǩ��ʱ�� = '"+dSet.Tables["CONFERENCE"].Rows[0][2].ToString()+"' WHERE (����ID = "+intConferenceID.ToString()+") AND (��ԱID = "+dataGridViewreSign.SelectedRows[i].Cells[0].Value.ToString()+")";
                sqlComm.ExecuteNonQuery();
               
            }

            sqlComm.CommandText = "SELECT ��Ա��.��ԱID, ��Ա��.����, ��Ա��.������λ, ��Ա��.ְ��, ��Ա��.ְ��, �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ��, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.���������� FROM ��Ա�� INNER JOIN �λ���Ա�� ON ��Ա��.��ԱID = �λ���Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ")  AND (�λ���Ա��.ǩ��ʱ�� IS NULL)";
            dSet.Tables["ReSIGN"].Clear();
            sqlDA.Fill(dSet, "ReSIGN");



            sqlConn.Close();

        }

        private void dataGridViewreSign_SelectionChanged(object sender, EventArgs e)
        {
            
            
            changeStatusBar();
        }
    }
}