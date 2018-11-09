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

            sqlComm.CommandText = "SELECT 会议主题, 开始时间, 结束时间 FROM 会议表 WHERE (会议ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONFERENCE");
            //dataGridViewreSign.DataSource = dSet.Tables["CONFERENCE"];

            sqlComm.CommandText = "SELECT 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            if (dSet.Tables.Contains("系统参数表")) dSet.Tables.Remove("系统参数表");
            sqlDA.Fill(dSet, "系统参数表");


            sqlComm.CommandText = "SELECT 人员表.人员ID, 人员表.姓名, 人员表.所属单位, 人员表.职称, 人员表.职务, 参会人员表.签到时间, 参会人员表.签出时间, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ")  AND (参会人员表.签到时间 IS NULL)";
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

            toolStripStatusLabelresign.Text = "共选择了" + dataGridViewreSign .SelectedRows.Count+ "人";

            if (dataGridViewreSign.SelectedRows.Count<1)
            {
                return;
            }
            
            int i;
            if (dSet.Tables["系统参数表"].Rows.Count > 0)
            {
                for (i = 0; i < 4; i++)
                {
                    if (dSet.Tables["系统参数表"].Rows[0][i].ToString() != "" && dataGridViewreSign.SelectedRows[0].Cells[7+i].Value.ToString()!= "")
                    {
                        labelAddi.Text = labelAddi.Text + dSet.Tables["系统参数表"].Rows[0][i].ToString() + ":" + dataGridViewreSign.SelectedRows[0].Cells[7 + i].Value.ToString() + "　";
                    }
                }
            }

        }

        private void btnreSign_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否要进行补签？该过程不可回退!", "信息提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                return;

            sqlConn.Open();
            int i;
            for(i=0;i<dataGridViewreSign.SelectedRows.Count;i++)
            {
                sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '"+dSet.Tables["CONFERENCE"].Rows[0][1].ToString()+"', 签出时间 = '"+dSet.Tables["CONFERENCE"].Rows[0][2].ToString()+"' WHERE (会议ID = "+intConferenceID.ToString()+") AND (人员ID = "+dataGridViewreSign.SelectedRows[i].Cells[0].Value.ToString()+")";
                sqlComm.ExecuteNonQuery();
               
            }

            sqlComm.CommandText = "SELECT 人员表.人员ID, 人员表.姓名, 人员表.所属单位, 人员表.职称, 人员表.职务, 参会人员表.签到时间, 参会人员表.签出时间, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四 FROM 人员表 INNER JOIN 参会人员表 ON 人员表.人员ID = 参会人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ")  AND (参会人员表.签到时间 IS NULL)";
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