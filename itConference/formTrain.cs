using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace itConference
{
    public partial class formTrain : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();
        public string strVersion = "0";
        private DataView dv;

        public string strConn = "";
        public int intMode = 0;
        public int intConferenceID = 0;

        public formTrain()
        {
            InitializeComponent();
        }

        private void formTrain_Load(object sender, EventArgs e)
        {
            int i;
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;

            sqlDA.SelectCommand = sqlComm;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT 单位名称 FROM 单位表 ORDER BY 上级单位";
            sqlDA.Fill(dSet, "所属单位");



            sqlConn.Close();

            comboBoxSDW.DataSource = dSet.Tables["所属单位"];
            comboBoxSDW.DisplayMember = "单位名称";
            comboBoxSDW.Text = "";

            switch (intMode)
            {
                case 1: //增加模式
                    sqlConn.Open();
                    sqlComm.CommandText = "SELECT 参会标记.参会, 单位表.ID, 单位表.单位名称, A.参会人数 FROM 单位表 LEFT OUTER JOIN (SELECT ID, 会议ID, 参会单位, 参会人数, 到会人数 FROM 培训人员表 WHERE (会议ID = 0)) A ON 单位表.单位名称 = A.参会单位 COLLATE Chinese_PRC_CI_AS CROSS JOIN 参会标记 ORDER BY 单位表.上级单位";
                    sqlDA.Fill(dSet, "单位表");

                    for (i = 0; i < dSet.Tables["单位表"].Rows.Count; i++)
                    {
                        if (dSet.Tables["单位表"].Rows[i][3].ToString() == "")
                            dSet.Tables["单位表"].Rows[i][3] = 0;
                    }

                    sqlConn.Close();

                    dv = dSet.Tables["单位表"].DefaultView;
                    dataGridViewPeople.DataSource = dv;
                    dataGridViewPeople.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                    dataGridViewPeople.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[1].Visible = false;
                    dataGridViewPeople.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

                    foreach (DataGridViewColumn c in dataGridViewPeople.Columns)
                    {
                        if (c.Index != 0) c.ReadOnly = true;
                    }

                    updateStateBar();
                    break;

                case 2:  //修改模式
                    
                    btnInput.Visible = false;

                    sqlConn.Open();

                    sqlComm.CommandText = "SELECT 会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间 FROM 培训表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
                    sqlDA.Fill(dSet, "Conference");



                    sqlComm.CommandText = "SELECT 参会标记.参会, 单位表.ID, 单位表.单位名称, A.参会人数 FROM 单位表 LEFT OUTER JOIN (SELECT ID, 会议ID, 参会单位, 参会人数, 到会人数 FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ")) A ON 单位表.单位名称 = A.参会单位 COLLATE Chinese_PRC_CI_AS CROSS JOIN 参会标记 ORDER BY 单位表.上级单位";
                    sqlDA.Fill(dSet, "单位表");

                    sqlConn.Close();

                    textBoxCname.Text = dSet.Tables["Conference"].Rows[0][0].ToString();
                    textBoxCname2.Text = dSet.Tables["Conference"].Rows[0][1].ToString();
                    textBoxConferenceNo.Text = dSet.Tables["Conference"].Rows[0][3].ToString();
                    textBoxCintro.Text = dSet.Tables["Conference"].Rows[0][2].ToString();
                    dateTimePickerStart.Value = Convert.ToDateTime(dSet.Tables["Conference"].Rows[0][4].ToString());
                    dateTimePickerEnd.Value = Convert.ToDateTime(dSet.Tables["Conference"].Rows[0][5].ToString());


                    foreach (DataRow r in dSet.Tables["单位表"].Rows)
                    {
                        if (r[3].ToString() != "")
                            r[0] = true;
                        else
                        {
                            r[0] = false;
                            r[3]=0;
                        }
                    }


                    dv = dSet.Tables["单位表"].DefaultView;
                    dataGridViewPeople.DataSource = dv;
                    dataGridViewPeople.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                    dataGridViewPeople.VirtualMode = true;
                    //dataGridViewPeople.Columns.Insert(0, new DataGridViewCheckBoxColumn());

                    dataGridViewPeople.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[1].Visible = false;
                    dataGridViewPeople.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

                    updateStateBar();
                    break;
                default:
                    break;
            }


        }
        private void updateStateBar()
        {
            int intCount = 0;

            for (int i = 0; i < dSet.Tables["单位表"].Rows.Count; i++)
            {
                if (Convert.ToBoolean(dSet.Tables["单位表"].Rows[i][0].ToString()))
                    intCount += (int)dSet.Tables["单位表"].Rows[i][3];
            }
            toolStripStatusLabelConference.Text = "参加会议的人数为：" + intCount.ToString() + "人";
        }

        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePickerEnd.Value < dateTimePickerStart.Value)
                dateTimePickerEnd.Value = dateTimePickerStart.Value.AddMinutes(10);
        }

        private void dataGridViewPeople_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                updateStateBar();
            }
        }

        private void btnSite_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgvP in dataGridViewPeople.SelectedRows)
            {
                dgvP.Cells[3].Value = numericUpDownCHRS.Value.ToString();
            }
            updateStateBar();
        }

        private void btnS_Click(object sender, EventArgs e)
        {
            DataView dv1 = new DataView(dSet.Tables["单位表"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;

            bool bFirst = true;
            string sSearch = "";

            if (comboBoxSDW.Text.Trim() != "")
            {
                sSearch = " 单位名称 LIKE '%" + comboBoxSDW.Text.Trim() + "%'";
                bFirst = false;
            }

            if (!bFirst)
                dv1.RowFilter = sSearch;
        }

        private void btnsAll_Click(object sender, EventArgs e)
        {
            DataView dv1 = new DataView(dSet.Tables["单位表"]);
            dataGridViewPeople.DataSource = dv1;
            dv.RowFilter = "";
        }

        private void btnCH_Click(object sender, EventArgs e)
        {
            DataView dv1 = new DataView(dSet.Tables["单位表"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;
            dv1.RowFilter = " (参会 = 1) ";
        }

        private void btnWCH_Click(object sender, EventArgs e)
        {
            DataView dv1 = new DataView(dSet.Tables["单位表"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;
            dv1.RowFilter = " (参会 = 0) ";
        }

        private void btnPSelect_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgvP in dataGridViewPeople.SelectedRows)
            {
                dgvP.Cells[0].Value = 1;
            }
            updateStateBar();
        }

        private void btnPSelectI_Click(object sender, EventArgs e)
        {
            int i;

            for (i = 0; i < dataGridViewPeople.Rows.Count; i++)
            {
                if ((bool)dataGridViewPeople.Rows[i].Cells[0].Value) dataGridViewPeople.Rows[i].Cells[0].Value = false;
                else dataGridViewPeople.Rows[i].Cells[0].Value = true;
            }
            updateStateBar();
        }

        private void btnPSelectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgvP in dataGridViewPeople.Rows)
            {
                dgvP.Cells[0].Value = true;
            }
            updateStateBar();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            int i;

            for (i = 0; i < dataGridViewPeople.Rows.Count; i++)
            {
                dataGridViewPeople.Rows[i].Cells[0].Value = 0;
            }
            updateStateBar();
        }

        private void btnYes_Click(object sender, EventArgs e)
        {

            System.Data.SqlClient.SqlTransaction sqlta;
            int i;

            if (textBoxCname.Text.Trim() == "")
            {
                MessageBox.Show("会议名称不能为空！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.Connection = sqlConn;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;


            switch (intMode)
            {
                case 1://增加会议

                    try
                    {
                        sqlComm.CommandText = "INSERT INTO 培训表 (会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间) VALUES (N'" + textBoxCname.Text.Trim() + "', N'" + textBoxCname2.Text.Trim() + "', N'" + textBoxCintro.Text.Trim() + "', N'" + textBoxConferenceNo.Text.Trim() + "', '" + dateTimePickerStart.Value.ToString() + "', '" + dateTimePickerEnd.Value.ToString() + "')";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "SELECT @@IDENTITY";
                        sqldr = sqlComm.ExecuteReader();
                        sqldr.Read();
                        string strCId = sqldr.GetValue(0).ToString();
                        sqldr.Close();


                        //参会人员表
                        btnsAll_Click(null, null);
                        foreach (DataGridViewRow r in dataGridViewPeople.Rows)
                        {
                            if ((bool)r.Cells[0].Value && (int)r.Cells[3].Value>0)
                            {
                                
                                sqlComm.CommandText = "INSERT INTO 培训人员表 (会议ID, 参会单位, 参会人数, 到会人数) VALUES (" + strCId + ", N'"+r.Cells[2].Value.ToString().Trim()+"', "+r.Cells[3].Value.ToString()+", 0)";
                                sqlComm.ExecuteNonQuery();
                            }
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
                        MessageBox.Show("会议已加入到数据库！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    sqlConn.Close();

                    break;
                case 2://修改会议
                    
                    try
                    {
                        sqlComm.CommandText = "UPDATE 培训表 SET 会议主题 = N'" + textBoxCname.Text.Trim() + "', 会议副题 = N'" + textBoxCname2.Text.Trim() + "', 会议内容 = N'" + textBoxCintro.Text.Trim() + "', 会场号 = N'" + textBoxConferenceNo.Text.Trim() + "', 开始时间 = '" + dateTimePickerStart.Value.ToString() + "', 结束时间 = '" + dateTimePickerEnd.Value.ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();



                        //参会人员表
                        sqlComm.CommandText = "DELETE FROM 培训签到表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "DELETE FROM 培训新到人员表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "DELETE FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();

                        btnsAll_Click(null, null);
                        foreach (DataGridViewRow r in dataGridViewPeople.Rows)
                        {
                            if ((bool)r.Cells[0].Value && (int)r.Cells[3].Value>0)
                            {

                                sqlComm.CommandText = "INSERT INTO 培训人员表 (会议ID, 参会单位, 参会人数, 到会人数) VALUES (" + intConferenceID.ToString() + ", N'" + r.Cells[2].Value.ToString().Trim() + "', " + r.Cells[3].Value.ToString() + ", 0)";
                                sqlComm.ExecuteNonQuery();
                            }
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
                        MessageBox.Show("会议信息已修改！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    sqlConn.Close();
                   
                    break;

            }

        }

        private void btnInput_Click(object sender, EventArgs e)
        {
            if (openFileDialogExcel.ShowDialog() == DialogResult.OK)
            {
                //string strOledbConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + openFileDialogExcel.FileName + "; 'Extended Properties=Excel 8.0;HDR=1;IMEX=1'";
                string strOledbConn = "provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + openFileDialogExcel.FileName + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";

                OleDbConnection oledbConn = new OleDbConnection(strOledbConn);

                oledbConn.Open();
                DataTable dt = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string tableName = dt.Rows[0][2].ToString().Trim();


                string strExcel = "";
                OleDbDataAdapter oledbDataAdapter = null;
                DataSet oledbdSet = new DataSet();
                strExcel = "select * from [" + tableName + "]";

                oledbDataAdapter = new OleDbDataAdapter(strExcel, oledbConn);
                oledbDataAdapter.Fill(oledbdSet, "会议表");
                oledbConn.Close();


                //导入数据
                System.Data.DataTable dtablePeople = oledbdSet.Tables["会议表"];
                btnClear_Click(null, null);
                textBoxCname2.Text = "";
                textBoxConferenceNo.Text = "";
                textBoxCintro.Text = "";



                int i, j;
                try
                {
                    for (i = 1; i < dtablePeople.Rows.Count; i++)
                    {
                        if (i == 1) textBoxCname.Text = dtablePeople.Rows[1][1].ToString();
                        //string sss = dtablePeople.Columns[0].ColumnName.ToString();
                        //string sss1 = "姓名='" + dtablePeople.Rows[i][3].ToString() + "' AND 所属单位='" + dtablePeople.Rows[i][2].ToString() + "'";

                        DataRow[] dr = dSet.Tables["单位表"].Select("单位名称='" + dtablePeople.Rows[i][2].ToString() + "'");

                        if (dr.Length > 0)
                        {
                            dr[0][0] = 1;
                            if(dtablePeople.Rows[i][3].ToString()=="")
                            {
                                dr[0][3] = 1;
                            }
                            else
                                dr[0][3] = Int32.Parse(dtablePeople.Rows[i][3].ToString());
                        }

                    }
                }
                catch
                {
                    sqlConn.Close();
                    MessageBox.Show("数据导入失败！", "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                }
                MessageBox.Show("数据导入成功！", "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                updateStateBar();

            }
            else
                return;
 
        }


    }
}
