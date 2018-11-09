using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class FormSelectCon : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";
        public string strVersion = "0";

        public int UserID = 0;

        public FormSelectCon()
        {
            InitializeComponent();
        }


        private void FormSelectCon_Load(object sender, EventArgs e)
        {
            if (UserID == 0)
            {
                this.Close();
                return;
            }
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlConn.Open();

            sqlComm.CommandText = "SELECT 参会标记.参会, 会议表.会议ID, 会议表.会议主题, 会议表.会议副题, 会议表.开始时间, 会议表.结束时间, 会议表.会场号 FROM 参会标记 CROSS JOIN 会议表";

            sqlDA.Fill(dSet, "会议表");
            sqlConn.Close();

            dataGridViewCon.DataSource = dSet.Tables["会议表"];
            dataGridViewCon.Columns[1].Visible = false;
            dataGridViewCon.Columns[2].ReadOnly = true;
            dataGridViewCon.Columns[3].ReadOnly = true;
            dataGridViewCon.Columns[4].ReadOnly = true;
            dataGridViewCon.Columns[5].ReadOnly = true;
            dataGridViewCon.Columns[6].ReadOnly = true;
        }

        private void buttonALL_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgvP in dataGridViewCon.Rows)
            {
                dgvP.Cells[0].Value = true;
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (UserID == 0)
                return;

            sqlConn.Open();

            System.Data.SqlClient.SqlTransaction sqlta;
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            try
            {
                 foreach (DataGridViewRow dgvP in dataGridViewCon.Rows)
                {
                    if(bool.Parse(dgvP.Cells[0].Value.ToString()))
                    {
                        sqlComm.CommandText = "INSERT INTO 参会人员表 (会议ID, 人员ID, 分组名称, 座位号) VALUES (" + dgvP.Cells[1].Value.ToString() + ", " + UserID.ToString() + ", NULL, NULL)";
                        sqlComm.ExecuteNonQuery();
                    }
                }

                sqlta.Commit();
            }
            catch (Exception err)
            {
                //回滚
                MessageBox.Show("数据库错误！", "数据库错误"+err.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("人员已加入到所选会议！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            sqlConn.Close();
            this.Close();
        }


    }
}
