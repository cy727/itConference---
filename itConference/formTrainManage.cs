using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formTrainManage : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();
        public string strVersion = "0";

        public string strConn = "";
        public int intMod = 0;


        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;

        public formTrainManage()
        {
            InitializeComponent();
        }

        private void formTrainManage_Load(object sender, EventArgs e)
        {
            
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            InitializeConferenceDataGridView();

            switch (intMod)
            {
                case 0://管理
                    btnAdd.Visible = true;
                    btnDel.Visible = true;
                    btnEdit.Visible = true;

                    btnSign.Visible = false;
                    btnOffsign.Visible = false;
                    btnResign.Visible = false;

                    btnCount.Visible = false;
                    //btnCountAll.Visible = false;

                    break;
                case 1://签到
                    btnAdd.Visible = false;
                    btnDel.Visible = false;
                    btnEdit.Visible = false;

                    btnSign.Visible = true;
                    btnOffsign.Visible = true;
                    btnResign.Visible = true;

                    btnCount.Visible = false;
                    //btnCountAll.Visible = false;
                    break;
                case 2://统计
                    btnAdd.Visible = false;
                    btnDel.Visible = false;
                    btnEdit.Visible = false;

                    btnSign.Visible = false;
                    btnOffsign.Visible = false;
                    btnResign.Visible = false;

                    btnCount.Visible = true;
                    //btnCountAll.Visible = true;
                    break;
                default:
                    btnAdd.Visible = false;
                    btnDel.Visible = false;
                    btnEdit.Visible = false;

                    btnSign.Visible = false;
                    btnOffsign.Visible = false;
                    btnResign.Visible = false;

                    btnCount.Visible = false;
                    //btnCountAll.Visible = false;

                    break;

            }

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            formTrain frmTrain = new formTrain();

            frmTrain.strConn = strConn;

            frmTrain.intMode = 1; //增加模式
            frmTrain.strVersion = strVersion;

            frmTrain.ShowDialog(this);

            //刷新会议列表
            InitializeConferenceDataGridView();

        }

        private void InitializeConferenceDataGridView()
        {

            sqlConn.Open();
            sqlComm.CommandText = "SELECT 会议ID, 会议主题, 会议副题, 开始时间 FROM 培训表 ORDER BY 开始时间 DESC";

            if (dSet.Tables.Contains("CONFERENCE")) dSet.Tables.Remove("CONFERENCE");
            sqlDA.Fill(dSet, "CONFERENCE");

            dataGridViewTrain.DataSource = dSet.Tables["CONFERENCE"];

            dataGridViewTrain.Columns[0].Visible = false;

            dataGridViewTrain.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewTrain.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewTrain.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewTrain.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            sqlConn.Close();

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

            if (dataGridViewTrain.SelectedRows.Count < 1)
                return;

            formTrain frmTrain = new formTrain();

            frmTrain.strConn = strConn;

            frmTrain.intMode = 2; //修改模式
            frmTrain.strVersion = strVersion;
            frmTrain.intConferenceID = Int32.Parse(dataGridViewTrain.SelectedRows[0].Cells[0].Value.ToString());


            frmTrain.ShowDialog(this);

            //刷新会议列表
            InitializeConferenceDataGridView();
        }

        private void dataGridViewTrain_DoubleClick(object sender, EventArgs e)
        {
            btnEdit_Click(null, null);
        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            if ((dataGridViewTrain.SelectedRows.Count == 0) || (dataGridViewTrain.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("请选择要删除的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("是否要删除的选定的会议信息？这个操作无法恢复！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //删除选定的会议
            sqlConn.Open();

            int i = 0;
            int intCID;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            try
            {
                for (i = 0; i < dataGridViewTrain.SelectedRows.Count; i++)
                {
                    intCID = Int32.Parse(dataGridViewTrain.SelectedRows[i].Cells[0].Value.ToString());
                    sqlComm.CommandText = "DELETE FROM 培训表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 培训人员表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 培训签到表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 培训新到人员表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                }
                sqlta.Commit();
            }
            catch (Exception err)
            {
                //回滚
                MessageBox.Show("数据库错误！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;
            }
            finally
            {
                MessageBox.Show("成功删除" + i.ToString() + "条会议信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            sqlConn.Close();

            //刷新会议列表
            InitializeConferenceDataGridView();

        }

        private void btnSign_Click(object sender, EventArgs e)
        {
            if (dataGridViewTrain.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formTrainSign frmTrainSign = new formTrainSign();

            frmTrainSign.strConn = strConn;
            frmTrainSign.strVersion = strVersion;
            frmTrainSign.intCardLength = intCardLength;
            frmTrainSign.intCard = intCard;
            frmTrainSign.intComNumber = intComNumber;
            frmTrainSign.lComBand = lComBand;

            frmTrainSign.intConferenceID = Int32.Parse(dataGridViewTrain.SelectedRows[0].Cells[0].Value.ToString());

            frmTrainSign.ShowDialog(this);

        }

        private void btnResign_Click(object sender, EventArgs e)
        {
            if (dataGridViewTrain.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formTrainResign frmTrainResign = new formTrainResign();

            frmTrainResign.strConn = strConn;
            //frmTrainResign.strVersion = strVersion;

            frmTrainResign.intConferenceID = Int32.Parse(dataGridViewTrain.SelectedRows[0].Cells[0].Value.ToString());

            frmTrainResign.ShowDialog(this);
        }

        private void btnOffsign_Click(object sender, EventArgs e)
        {
            formoffSign frmoffSign = new formoffSign();
            frmoffSign.strVersion = strVersion;
            frmoffSign.intCardLength = intCardLength;
            frmoffSign.intCard = intCard;
            frmoffSign.intComNumber = intComNumber;
            frmoffSign.lComBand = lComBand;
            frmoffSign.Show(this);
        }

        private void btnCount_Click(object sender, EventArgs e)
        {

            if (dataGridViewTrain.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formTarinCount frmTarinCount = new formTarinCount();

            frmTarinCount.strConn = strConn;
            frmTarinCount.strVersion = strVersion;
            frmTarinCount.intConferenceID = Int32.Parse(dataGridViewTrain.SelectedRows[0].Cells[0].Value.ToString());

            frmTarinCount.ShowDialog(this);
        }
    }
}
