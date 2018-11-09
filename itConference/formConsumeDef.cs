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

            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                case "2": //标准版
                    MessageBox.Show("该版本不支持此项功能！请购买高级版本。", "版本信息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "3": //豪华版
                    break;
                default:
                    break;
            } 
            
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT 会议主题, 开始时间, 结束时间 FROM 会议表 WHERE (会议ID = "+intConferenceID.ToString()+")";
            sqlDA.Fill(dSet, "CONFERENCE");

            labelConference.Text = dSet.Tables["CONFERENCE"].Rows[0][0].ToString();
            dateTimePickerCstart.Value = DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][1].ToString());
            dateTimePickerCend.Value = DateTime.Parse(dSet.Tables["CONFERENCE"].Rows[0][2].ToString());

            sqlComm.CommandText = "SELECT 消费ID, 消费名称, 开始时间, 结束时间 FROM 会议消费表 WHERE (会议ID = "+intConferenceID.ToString()+")";
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
                MessageBox.Show("会议消费名称不能为空！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.CommandText = "INSERT INTO 会议消费表 (消费名称, 开始时间, 结束时间, 会议ID) VALUES (N'" + textBoxCname.Text.Trim() + "', '" + dateTimePickerCstart.Value.ToString() + "', '" + dateTimePickerCend.Value.ToString() + "', "+intConferenceID.ToString()+")";
            sqlComm.ExecuteNonQuery();

            if (dSet.Tables.Contains("CONSUME")) dSet.Tables.Remove("CONSUME");
            sqlComm.CommandText = "SELECT 消费ID, 消费名称, 开始时间, 结束时间 FROM 会议消费表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "CONSUME");

            dataGridViewConsume.DataSource = dSet.Tables["CONSUME"];
            dataGridViewConsume.Columns[0].Visible = false;
            sqlConn.Close();

        }

        private void btncEdit_Click(object sender, EventArgs e)
        {
            if (textBoxCname.Text.Trim() == "")
            {
                MessageBox.Show("会议消费名称不能为空！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.CommandText = "UPDATE 会议消费表 SET 消费名称 = N'" + textBoxCname.Text.Trim() + "', 开始时间 = '" + dateTimePickerCstart.Value.ToString() + "', 结束时间 = '" + dateTimePickerCend.Value.ToString() + "' WHERE (消费ID = " + dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString()+ ")";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();

            dataGridViewConsume.SelectedRows[0].Cells[1].Value = textBoxCname.Text.Trim();
            dataGridViewConsume.SelectedRows[0].Cells[2].Value = dateTimePickerCstart.Value;
            dataGridViewConsume.SelectedRows[0].Cells[3].Value = dateTimePickerCend.Value;

        }

        private void btncDel_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("是否删除该项会议消费？该过程不可回退。", "输入", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.No) return;
            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM 会议消费表 WHERE (消费ID = " + dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString() + ")";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();

            dataGridViewConsume.Rows.Remove(dataGridViewConsume.SelectedRows[0]);
        }

        private void textBoxCname_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCname.Text.Length > 100)
            {
                this.errorProviderC.SetError(this.textBoxCname, "字段过长");
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