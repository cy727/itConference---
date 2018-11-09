using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using SmsATCommdll;

namespace itConference
{
    public partial class formConference : Form
    {
        public string strConn;
        public int intMode, intConferenceID;
        private System.Data.DataSet dSet = new DataSet();

        //private System.Collections.Generic.Dictionary<int, bool> pCheckState;
        public string strVersion="0";
        private DataView dv;
        SmsATComm at = new SmsATComm();

        public int intCount = 0; //参会人数
        public int intWTZ = 0; //未通知

        public formConference()
        {
            InitializeComponent();
        }

        private void formConference_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            int i;

            //版本处理
            switch(strVersion)
            {
                case "1": //基础版
                    btnSite.Enabled = false;
                    btnGroup.Enabled = false;
                    break;
                case "2": //标准版
                    btnSite.Enabled = false;
                    break;
                case "3": //豪华版
                    break;
                default:
                    //btnYes.Enabled = false;
                    break;
            }

            sqlDA.SelectCommand = sqlComm;
            //pCheckState = new Dictionary<int, bool>();
            dateTimePickerEnd.Value = dateTimePickerStart.Value.AddHours(1);

            sqlConn.Open();
            sqlComm.CommandText = "SELECT 分组名称 FROM 人员分组表";
            sqlDA.Fill(dSet, "人员分组表");

            sqlComm.CommandText = "SELECT DISTINCT 所属单位 FROM 人员表 WHERE (所属单位 <> N'') ORDER BY 所属单位";
            sqlDA.Fill(dSet, "所属单位");

            sqlComm.CommandText = "SELECT DISTINCT 职务 FROM 人员表 WHERE (职务 <> N'') ORDER BY 职务";
            sqlDA.Fill(dSet, "职务");
            sqlConn.Close();

            for(i=0;i<dSet.Tables["人员分组表"].Rows.Count;i++)
            {
                comboBoxGroup.Items.Add(dSet.Tables["人员分组表"].Rows[i][0].ToString());
            }
            comboBoxSDW.DataSource = dSet.Tables["所属单位"];
            comboBoxSDW.DisplayMember = "所属单位";
            comboBoxSDW.Text = "";
            comboBoxSZW.DataSource = dSet.Tables["职务"];
            comboBoxSZW.DisplayMember = "职务";
            comboBoxSZW.Text = "";

            switch (intMode)
            {
                case 1: //增加模式
                    sqlConn.Open();
                    sqlComm.CommandText = "SELECT 参会标记.参会, 人员表.人员ID, 人员表.姓名, 人员表.所属单位, 人员表.人员类别, A.分组名称, A.座位号, 人员表.编号, 人员表.职务, 人员表.手机, 人员表.EMail, A.短信, A.邮件, A.回复 FROM 人员表 LEFT OUTER JOIN (SELECT 人员ID AS RID, 会议ID, 分组名称, 座位号, 短信, 邮件, 回复 FROM 参会人员表 WHERE (会议ID = 0)) A ON 人员表.人员ID = A.RID CROSS JOIN 参会标记";
                    sqlDA.Fill(dSet, "People");

                    sqlConn.Close();

                    dv = dSet.Tables["People"].DefaultView;
                    dataGridViewPeople.DataSource = dv;
                    dataGridViewPeople.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                    /*
                    dataGridViewPeople.VirtualMode = true;
                    dataGridViewPeople.Columns.Insert(0, new DataGridViewCheckBoxColumn());
                    dataGridViewPeople.Columns[0].Resizable = DataGridViewTriState.False;
                    dataGridViewPeople.Columns[0].Frozen = true;
                    dataGridViewPeople.Columns[0].DividerWidth = 1;
                    dataGridViewPeople.Columns[0].Width = 30;
                     */
                    dataGridViewPeople.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[1].Visible=false;
                    dataGridViewPeople.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[10].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[12].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[13].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;


                    foreach (DataGridViewColumn c in dataGridViewPeople.Columns)
                    {
                        if (c.Index != 0) c.ReadOnly = true;
                    }
                    /*
 
foreach (DataGridViewRow r in dataGridViewPeople.Rows)
{
    r.Cells[3].Value = ""; 
    r.Cells[5].Value = "";
}
 */
                    
                    updateStateBar();
                    break;

                case 2:  //修改模式
                    btnInput.Visible = false;
                    sqlConn.Open();

                    sqlComm.CommandText="SELECT 会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间 FROM 会议表 WHERE (会议ID = "+intConferenceID.ToString()+")";
                    sqlDA.Fill(dSet, "Conference");


                    sqlComm.CommandText = "SELECT 分组名称 FROM 分组表 WHERE (会议ID = "+intConferenceID.ToString()+")";
                    sqlDA.Fill(dSet, "Group");

                    sqlComm.CommandText = "SELECT  参会标记.参会,人员表.人员ID, 人员表.姓名, 人员表.所属单位, 人员表.人员类别, A.分组名称, A.座位号, A.会议ID, 人员表.编号, 人员表.职务, 人员表.手机, 人员表.EMail, A.短信, A.邮件, A.回复 FROM 人员表 LEFT OUTER JOIN (SELECT 人员ID AS RID, 会议ID, 分组名称, 座位号, 短信, 邮件, 回复 FROM 参会人员表 WHERE (会议ID = " + intConferenceID.ToString() + ")) A ON 人员表.人员ID = A.RID CROSS JOIN 参会标记";
                    sqlDA.Fill(dSet, "People");

                    sqlComm.CommandText = "SELECT 人员ID FROM 参会人员表 WHERE (会议ID = "+intConferenceID.ToString()+")";
                    sqlDA.Fill(dSet, "cPeople");//参会人员列表

                    sqlConn.Close();

                    textBoxCname.Text = dSet.Tables["Conference"].Rows[0][0].ToString();
                    textBoxCname2.Text = dSet.Tables["Conference"].Rows[0][1].ToString();
                    textBoxConferenceNo.Text = dSet.Tables["Conference"].Rows[0][3].ToString();
                    textBoxCintro.Text = dSet.Tables["Conference"].Rows[0][2].ToString();
                    dateTimePickerStart.Value = Convert.ToDateTime(dSet.Tables["Conference"].Rows[0][4].ToString());
                    dateTimePickerEnd.Value = Convert.ToDateTime(dSet.Tables["Conference"].Rows[0][5].ToString());

                    for (i = 0; i < dSet.Tables["Group"].Rows.Count; i++)
                    {
                        listBoxGroup.Items.Add(dSet.Tables["Group"].Rows[i][0]);
                    }

                    foreach (DataRow r in dSet.Tables["People"].Rows)
                    {
                        if (r[7].ToString() == intConferenceID.ToString())
                            r[0] = true;
                        else
                            r[0] = false;
                    }
                    dSet.Tables["People"].Columns.Remove("会议ID");

                    dv = dSet.Tables["People"].DefaultView;
                    dataGridViewPeople.DataSource = dv;
                    dataGridViewPeople.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                    dataGridViewPeople.VirtualMode = true;
                    //dataGridViewPeople.Columns.Insert(0, new DataGridViewCheckBoxColumn());

                    dataGridViewPeople.Columns[7].Visible = false;

                    /*
                    dataGridViewPeople.Columns[0].Resizable = DataGridViewTriState.False;
                    dataGridViewPeople.Columns[0].Frozen = true;
                    dataGridViewPeople.Columns[0].DividerWidth = 1;
                    dataGridViewPeople.Columns[0].Width = 30;
                     */
                    dataGridViewPeople.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[1].Visible = false;
                    dataGridViewPeople.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[10].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[12].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridViewPeople.Columns[13].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;



                    updateStateBar();
                    break;
                default:
                    break;
            }
        }

        private void updateStateBar()
        {
            intCount = 0;
            intWTZ = 0;
            /*
            foreach (bool isChecked in pCheckState.Values)
            {
                if (isChecked) intCount++;
            }
             */
            for (int i = 0; i < dSet.Tables["People"].Rows.Count; i++)
            {
                if (Convert.ToBoolean(dSet.Tables["People"].Rows[i][0].ToString()))
                {
                    intCount++;

                    if (dSet.Tables["People"].Rows[i][11].ToString() == "")
                        intWTZ++;
                }
            }
            toolStripStatusLabelConference.Text = "参加会议的人数为：" + intCount.ToString() + "人";
        }

        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePickerEnd.Value < dateTimePickerStart.Value)
                dateTimePickerEnd.Value = dateTimePickerStart.Value.AddMinutes(10);
        }

        private void btnGAdd_Click(object sender, EventArgs e)
        {
            if(textBoxGroup.Text.Trim()=="") return;
            listBoxGroup.Items.Add(textBoxGroup.Text.Trim());
        }

        private void btnGEdit_Click(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim() == "") return;
            if (listBoxGroup.SelectedIndices.Count < 1) return;

            //修改列表分组
            foreach (DataGridViewRow r in dataGridViewPeople.Rows)
            {
                if (r.Cells[5].Value.ToString() == listBoxGroup.Items[listBoxGroup.SelectedIndices[0]].ToString())
                    r.Cells[5].Value = textBoxGroup.Text.Trim(); 
            }
            listBoxGroup.Items[listBoxGroup.SelectedIndices[0]] = textBoxGroup.Text.Trim();
        }

        private void btnGDel_Click(object sender, EventArgs e)
        {
            if (listBoxGroup.SelectedIndices.Count < 1) return;
            //修改列表
            foreach (DataGridViewRow r in dataGridViewPeople.Rows)
            {
                if (r.Cells[5].Value.ToString() == listBoxGroup.Items[listBoxGroup.SelectedIndices[0]].ToString())
                    r.Cells[5].Value = "";
            }
            listBoxGroup.Items.RemoveAt(listBoxGroup.SelectedIndices[0]);
        }

        private void dataGridViewPeople_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
            if (e.ColumnIndex == 0)
            {
                //dataGridViewPeople.Rows[e.RowIndex].Cells[0].Value = (bool)dataGridViewPeople.Rows[e.RowIndex].Cells[0].EditedFormattedValue;
                //updateStateBar();
            }
            
            
        }

        private void dataGridViewPeople_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                //dataGridViewPeople.Rows[e.RowIndex].Cells[0].Value = (bool)dataGridViewPeople.Rows[e.RowIndex].Cells[0].EditedFormattedValue;
                updateStateBar();
            }
        }

        private void dataGridViewPeople_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            
         
        }

        private void dataGridViewPeople_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {
            /*
            if (e.ColumnIndex == 0)
            {
                int firstColumn = (int)dataGridViewPeople.Rows[e.RowIndex].Cells[1].Value;
                if (!pCheckState.ContainsKey(firstColumn))
                {
                    pCheckState.Add(firstColumn,(bool)e.Value);
                }
                else
                {
                    pCheckState[firstColumn] = (bool)e.Value;
                }

            }
             * */
        }

        private void btnPSelectAll_Click(object sender, EventArgs e)
        {
            this.dataGridViewPeople.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);
            foreach (DataGridViewRow dgvP in dataGridViewPeople.Rows)
            {
               dgvP.Cells[0].Value = true;
            }
            updateStateBar();
            this.dataGridViewPeople.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            int i;
            this.dataGridViewPeople.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);


            for (i = 0; i < dataGridViewPeople.Rows.Count ; i++)
            {
                dataGridViewPeople.Rows[i].Cells[0].Value = 0;
            }
            updateStateBar();
            this.dataGridViewPeople.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);

        }

        private void btnPSelectI_Click(object sender, EventArgs e)
        {
            int i;
            this.dataGridViewPeople.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);

            for (i = 0; i < dataGridViewPeople.Rows.Count ; i++)
            {
                if((bool)dataGridViewPeople.Rows[i].Cells[0].Value) dataGridViewPeople.Rows[i].Cells[0].Value=false;
                else dataGridViewPeople.Rows[i].Cells[0].Value = true;
            }
            updateStateBar();
            this.dataGridViewPeople.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);

        }

        private void btnPSelect_Click(object sender, EventArgs e)
        {
            this.dataGridViewPeople.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);
            foreach (DataGridViewRow dgvP in dataGridViewPeople.SelectedRows)
            {
                dgvP.Cells[0].Value = 1;
            }
            updateStateBar();
            this.dataGridViewPeople.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);
        }

        private void btnGroup_Click(object sender, EventArgs e)
        {
            //if (textBoxGroup.Text.Trim() == "") return;

            if (listBoxGroup.SelectedIndices.Count < 1) return;
            foreach (DataGridViewRow dgvP in dataGridViewPeople.SelectedRows)
            {
                dgvP.Cells[5].Value = listBoxGroup.Items[listBoxGroup.SelectedIndices[0]];
            }
        }

        private void btnGClear_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgvP in dataGridViewPeople.SelectedRows)
            {
                dgvP.Cells[5].Value = "";
            }
        }

        
        private void btnYes_Click(object sender, EventArgs e)
        {

            if (intConferenceID != 0 && intMode == 1)
            {
                if (MessageBox.Show("会议已经保存，是否增加新的会议？", "会议管理", MessageBoxButtons.YesNo, MessageBoxIcon.Error) != DialogResult.Yes)
                    return;
            }


            System.Data.SqlClient.SqlDataReader sqldr;
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


            switch(intMode)
            {
                case 1://增加会议


                    try
                    {
                        sqlComm.CommandText = "INSERT INTO 会议表 (会议主题, 会议副题, 会议内容, 会场号, 开始时间, 结束时间) VALUES (N'" + textBoxCname.Text.Trim() + "', N'" + textBoxCname2.Text.Trim() + "', N'" + textBoxCintro.Text.Trim() + "', N'" + textBoxConferenceNo.Text.Trim() + "', '" + dateTimePickerStart.Value.ToString() + "', '" + dateTimePickerEnd .Value.ToString()+ "')";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "SELECT @@IDENTITY";
                        sqldr = sqlComm.ExecuteReader();
                        sqldr.Read();
                        string strCId = sqldr.GetValue(0).ToString();
                        intConferenceID = int.Parse(strCId);
                        sqldr.Close();


                        //分组表
                        for (i = 0; i < listBoxGroup.Items.Count; i++)
                        {
                            sqlComm.CommandText = "INSERT INTO 分组表 (会议ID, 分组名称) VALUES ("+strCId+", N'"+listBoxGroup.Items[i].ToString()+"')";
                            sqlComm.ExecuteNonQuery();
                        }

                        //参会人员表
                        btnsAll_Click(null, null);
                        foreach(DataGridViewRow r in dataGridViewPeople.Rows)
                        {
                            if ((bool)r.Cells[0].Value)
                            {
                                sqlComm.CommandText = "INSERT INTO 参会人员表 (会议ID, 人员ID, 分组名称, 座位号) VALUES (" + strCId + ", " + r.Cells[1].Value + ", N'" + r.Cells[5].Value + "', N'" + r.Cells[6].Value + "')";
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
                        sqlComm.CommandText = "UPDATE 会议表 SET 会议主题 = N'" + textBoxCname.Text.Trim() + "', 会议副题 = N'" + textBoxCname2.Text.Trim() + "', 会议内容 = N'" + textBoxCintro.Text.Trim() + "', 会场号 = N'" + textBoxConferenceNo.Text.Trim() + "', 开始时间 = '" + dateTimePickerStart.Value.ToString() + "', 结束时间 = '" + dateTimePickerEnd .Value.ToString()+ "' WHERE (会议ID = "+intConferenceID.ToString()+")";
                        sqlComm.ExecuteNonQuery();



                        //分组表
                        sqlComm.CommandText = "DELETE FROM 分组表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();
                        for (i = 0; i < listBoxGroup.Items.Count; i++)
                        {
                            sqlComm.CommandText = "INSERT INTO 分组表 (会议ID, 分组名称) VALUES (" + intConferenceID.ToString() + ", N'" + listBoxGroup.Items[i].ToString() + "')";
                            sqlComm.ExecuteNonQuery();
                        }

                        //参会人员表
                        System.Data.DataRow[] pRow;
                        //sqlComm.CommandText = "DELETE FROM 参会人员表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
                        //sqlComm.ExecuteNonQuery();
                        btnsAll_Click(null, null);
                        foreach (DataGridViewRow r in dataGridViewPeople.Rows)
                        {
                            pRow = dSet.Tables["cPeople"].Select("人员ID = '" + r.Cells[1].Value.ToString() + "'");
                            if ((bool)r.Cells[0].Value) //选择该人员
                            {
                                if (pRow.Length < 1) //没有该人员数据
                                {
                                    sqlComm.CommandText = "INSERT INTO 参会人员表 (会议ID, 人员ID, 分组名称, 座位号) VALUES (" + intConferenceID.ToString() + ", " + r.Cells[1].Value + ", N'" + r.Cells[5].Value + "', N'" + r.Cells[6].Value + "')";
                                    
                                }
                                else
                                {
                                    sqlComm.CommandText = "UPDATE 参会人员表 SET 分组名称 = N'" + r.Cells[5].Value + "', 座位号 = N'" + r.Cells[6].Value + "' WHERE (人员ID = " + r.Cells[1].Value + ") AND (会议ID = "+intConferenceID.ToString()+")";
                                }
                                sqlComm.ExecuteNonQuery();
                            }
                            else //未选该人员
                            {
                                if (pRow.Length > 0)
                                {
                                    sqlComm.CommandText = "DELETE FROM 参会人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID = " + r.Cells[1].Value.ToString() + ")";
                                    sqlComm.ExecuteNonQuery();
                                }
                                
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

        private void btnSite_Click(object sender, EventArgs e)
        {

            foreach (DataGridViewRow dgvP in dataGridViewPeople.SelectedRows)
            {
                dgvP.Cells[6].Value = textBoxSite.Text.Trim();
            }
        }

        private void dataGridViewPeople_Click(object sender, EventArgs e)
        {
            if (dataGridViewPeople.SelectedRows.Count < 1)
                return;

            textBoxSite.Text = dataGridViewPeople.SelectedRows[0].Cells[6].Value.ToString();
            comboBoxGroup.Text = dataGridViewPeople.SelectedRows[0].Cells[4].Value.ToString();
            
        }

        private void btnGSelect_Click(object sender, EventArgs e)
        {
            if (comboBoxGroup.Text == "") return;
            foreach (DataGridViewRow dgvP in dataGridViewPeople.Rows)
            {
                if (dgvP.Cells[4].Value.ToString()==comboBoxGroup.Text)
                    dgvP.Cells[0].Value = true;
            }
            updateStateBar();
        }

        private void btnGUnselect_Click(object sender, EventArgs e)
        {
            if (comboBoxGroup.Text == "") return;
            foreach (DataGridViewRow dgvP in dataGridViewPeople.Rows)
            {
                if (dgvP.Cells[4].Value.ToString() == comboBoxGroup.Text)
                    dgvP.Cells[0].Value = false;
            }
            updateStateBar();
        }

        private void textBoxCname_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCname.Text.Length > 100)
            {
                this.errorProviderConference.SetError(this.textBoxCname, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }

        private void textBoxCname2_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCname2.Text.Length > 100)
            {
                this.errorProviderConference.SetError(this.textBoxCname2, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }

        private void textBoxConferenceNo_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxConferenceNo.Text.Length > 50)
            {
                this.errorProviderConference.SetError(this.textBoxConferenceNo, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }

        private void textBoxCintro_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxConferenceNo.Text.Length > 200)
            {
                this.errorProviderConference.SetError(this.textBoxConferenceNo, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }

        private void comboBoxGroup_Validating(object sender, CancelEventArgs e)
        {
            if (this.comboBoxGroup.Text.Length > 100)
            {
                this.errorProviderConference.SetError(this.textBoxConferenceNo, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }

        private void textBoxSite_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxSite.Text.Length > 50)
            {
                this.errorProviderConference.SetError(this.textBoxSite, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }


        private void textBoxGroup_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxGroup.Text.Length > 100)
            {
                this.errorProviderConference.SetError(this.textBoxGroup, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderConference.Clear();
            }
        }

        private void listBoxGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxGroup.SelectedItems.Count < 1) return;

            textBoxGroup.Text = listBoxGroup.Items[listBoxGroup.SelectedIndices[0]].ToString();
        }

        private void btnS_Click(object sender, EventArgs e)
        {
            DataView dv1 = new DataView(dSet.Tables["People"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;

            bool bFirst = true;
            string sSearch = "";

            if (textBoxSXM.Text.Trim() != "")
            {
                sSearch = " 姓名 LIKE '%" + textBoxSXM.Text.Trim() + "%'";
                bFirst = false;
            }
            if (comboBoxSDW.Text.Trim() != "")
            {
                if (bFirst)
                    sSearch = " 所属单位 LIKE '%" + comboBoxSDW.Text.Trim() + "%'";
                else
                {
                    sSearch = sSearch + " AND 所属单位 LIKE '%" + comboBoxSDW.Text.Trim() + "%'";
                }
                bFirst = false;
            }
            if (textBoxXBH.Text.Trim() != "")
            {
                if (bFirst)
                    sSearch = " 编号 LIKE '%" + textBoxXBH.Text.Trim() + "%'";
                else
                {
                    sSearch = sSearch + " AND 编号 LIKE '%" + textBoxXBH.Text.Trim() + "%'";
                }
                bFirst = false;
            }
            if (comboBoxSZW.Text.Trim() != "")
            {
                if (bFirst)
                    sSearch = " 职务 LIKE '%" + comboBoxSZW.Text.Trim() + "%'";
                else
                {
                    sSearch = sSearch + " AND 职务 LIKE '%" + comboBoxSZW.Text.Trim() + "%'";
                }
                bFirst = false;
            }

            if (!bFirst)
                dv1.RowFilter = sSearch;
        }

        private void btnsAll_Click(object sender, EventArgs e)
        {
            DataView dv1 = new DataView(dSet.Tables["People"]);
            dataGridViewPeople.DataSource = dv1;
            dv.RowFilter = "";
        }

        private void btnCH_Click(object sender, EventArgs e)
        {
            DataView dv1=new DataView(dSet.Tables["People"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;
            dv1.RowFilter = " (参会 = 1) ";

        }

        private void btnWCH_Click(object sender, EventArgs e)
        {

            DataView dv1 = new DataView(dSet.Tables["People"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;
            dv1.RowFilter = " (参会 = 0) ";

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
                textBoxCname2.Text="";
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

                        DataRow[] dr = dSet.Tables["People"].Select("姓名='" + dtablePeople.Rows[i][3].ToString() + "' AND 所属单位='" + dtablePeople.Rows[i][2].ToString() + "'");

                        if (dr.Length > 0)
                        {
                            dr[0][0] = 1;
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

        private void btnTZ_Click(object sender, EventArgs e)
        {
            if (intConferenceID == 0)
            {
                MessageBox.Show("请先点击确定按钮保存会议！", "会议通知", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            formTZ frmTZ = new formTZ(this);
            frmTZ.ShowDialog(this);
        }

        void at_DataReceivedChanged(object sender, EventArgs e)
        {
            DataReceived();
        }

        private void DataReceived()
        {
            string sms = at.NewSms;
            string id = at.callID;
            if (sms.Length != 0)
            {
                MessageBox.Show("新信息:" + sms, "提示！");
                sms = "";
            }
            if (id.Length != 0)
            {
                if (DialogResult.OK == MessageBox.Show(id + "来电", "提示！", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk))
                {
                    at.AnswerPhone();
                    id = "";
                }
                else
                {
                    at.HangPhone();
                    id = "";
                }
            }

        }
    }
}