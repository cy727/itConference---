using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formTrainResign : Form
    {
        public string strConn;
        public int intConferenceID=0;
        private System.Data.DataSet dSet = new DataSet();
        
        public formTrainResign()
        {
            InitializeComponent();
        }

        private void formTrainResign_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;


            sqlDA.SelectCommand = sqlComm;
            sqlConn.Open();

            sqlComm.CommandText = "SELECT 会议主题, 开始时间, 结束时间 FROM 培训表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "CONFERENCE");
            //dataGridViewreSign.DataSource = dSet.Tables["CONFERENCE"];


            sqlComm.CommandText = "SELECT 参会单位, 参会人数 - 到会人数 AS 缺席人数, 0 AS 补签 FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会人数 - 到会人数 <> 0)";
            sqlDA.Fill(dSet, "ReSIGN");
            dataGridViewreSign.DataSource = dSet.Tables["ReSIGN"];
            dataGridViewreSign.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewreSign.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewreSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            labelConferencename.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();

            changeStatusBar();
            sqlConn.Close();
        }

        private void changeStatusBar()
        {

            
            toolStripStatusLabelresign.Text = "共选择了" + dataGridViewreSign.SelectedRows.Count + "个单位";

            if (dataGridViewreSign.SelectedRows.Count < 1)
            {
                return;
            }

       }

        private void btnBQ_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridViewreSign.SelectedRows.Count; i++)
            {
                if ((int)numericUpDownBQ.Value < (int)dataGridViewreSign.SelectedRows[i].Cells[1].Value)
                    dataGridViewreSign.SelectedRows[i].Cells[2].Value = numericUpDownBQ.Value;
                else
                    dataGridViewreSign.SelectedRows[i].Cells[2].Value = dataGridViewreSign.SelectedRows[i].Cells[1].Value;
            }
        }

        private void btnQB_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridViewreSign.Rows.Count; i++)
            {
                dataGridViewreSign.Rows[i].Cells[2].Value = dataGridViewreSign.Rows[i].Cells[1].Value;
            }
        }

        private void btnreSign_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否要进行补签？该过程不可回退!", "信息提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                return;

            sqlConn.Open();
            int i,j;
            for (i = 0; i < dataGridViewreSign.Rows.Count; i++)
            {
                if((int)dataGridViewreSign.Rows [i].Cells[2].Value<=0)
                    continue;

                sqlComm.CommandText = "UPDATE 培训人员表 SET 到会人数 = 到会人数 + " + dataGridViewreSign.Rows [i].Cells[2].Value.ToString()+ " WHERE (会议ID = "+intConferenceID.ToString()+") AND (参会单位 = N'"+dataGridViewreSign.Rows[i].Cells[0].Value.ToString()+"')";
                sqlComm.ExecuteNonQuery();

                for (j = 0; j < (int)dataGridViewreSign.Rows[i].Cells[2].Value; j++)
                {
                    sqlComm.CommandText = "INSERT INTO 培训签到表 (会议ID, 人员ID, 代表单位, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", 0, N'" + dataGridViewreSign.Rows[i].Cells[0].Value.ToString() + "', '" + dSet.Tables["CONFERENCE"].Rows[0][1].ToString() + "', '" + dSet.Tables["CONFERENCE"].Rows[0][2].ToString() + "')";
                    sqlComm.ExecuteNonQuery();
                }
            }

            sqlComm.CommandText = "SELECT 参会单位, 参会人数 - 到会人数 AS 缺席人数, 0 AS 补签 FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会人数 - 到会人数 <> 0)";
            dSet.Tables["ReSIGN"].Clear();
            sqlDA.Fill(dSet, "ReSIGN");


            sqlConn.Close();
        }

    }
}
