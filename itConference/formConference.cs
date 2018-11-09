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

        public int intCount = 0; //�λ�����
        public int intWTZ = 0; //δ֪ͨ

        public formConference()
        {
            InitializeComponent();
        }

        private void formConference_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            int i;

            //�汾����
            switch(strVersion)
            {
                case "1": //������
                    btnSite.Enabled = false;
                    btnGroup.Enabled = false;
                    break;
                case "2": //��׼��
                    btnSite.Enabled = false;
                    break;
                case "3": //������
                    break;
                default:
                    //btnYes.Enabled = false;
                    break;
            }

            sqlDA.SelectCommand = sqlComm;
            //pCheckState = new Dictionary<int, bool>();
            dateTimePickerEnd.Value = dateTimePickerStart.Value.AddHours(1);

            sqlConn.Open();
            sqlComm.CommandText = "SELECT �������� FROM ��Ա�����";
            sqlDA.Fill(dSet, "��Ա�����");

            sqlComm.CommandText = "SELECT DISTINCT ������λ FROM ��Ա�� WHERE (������λ <> N'') ORDER BY ������λ";
            sqlDA.Fill(dSet, "������λ");

            sqlComm.CommandText = "SELECT DISTINCT ְ�� FROM ��Ա�� WHERE (ְ�� <> N'') ORDER BY ְ��";
            sqlDA.Fill(dSet, "ְ��");
            sqlConn.Close();

            for(i=0;i<dSet.Tables["��Ա�����"].Rows.Count;i++)
            {
                comboBoxGroup.Items.Add(dSet.Tables["��Ա�����"].Rows[i][0].ToString());
            }
            comboBoxSDW.DataSource = dSet.Tables["������λ"];
            comboBoxSDW.DisplayMember = "������λ";
            comboBoxSDW.Text = "";
            comboBoxSZW.DataSource = dSet.Tables["ְ��"];
            comboBoxSZW.DisplayMember = "ְ��";
            comboBoxSZW.Text = "";

            switch (intMode)
            {
                case 1: //����ģʽ
                    sqlConn.Open();
                    sqlComm.CommandText = "SELECT �λ���.�λ�, ��Ա��.��ԱID, ��Ա��.����, ��Ա��.������λ, ��Ա��.��Ա���, A.��������, A.��λ��, ��Ա��.���, ��Ա��.ְ��, ��Ա��.�ֻ�, ��Ա��.EMail, A.����, A.�ʼ�, A.�ظ� FROM ��Ա�� LEFT OUTER JOIN (SELECT ��ԱID AS RID, ����ID, ��������, ��λ��, ����, �ʼ�, �ظ� FROM �λ���Ա�� WHERE (����ID = 0)) A ON ��Ա��.��ԱID = A.RID CROSS JOIN �λ���";
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

                case 2:  //�޸�ģʽ
                    btnInput.Visible = false;
                    sqlConn.Open();

                    sqlComm.CommandText="SELECT ��������, ���鸱��, ��������, �᳡��, ��ʼʱ��, ����ʱ�� FROM ����� WHERE (����ID = "+intConferenceID.ToString()+")";
                    sqlDA.Fill(dSet, "Conference");


                    sqlComm.CommandText = "SELECT �������� FROM ����� WHERE (����ID = "+intConferenceID.ToString()+")";
                    sqlDA.Fill(dSet, "Group");

                    sqlComm.CommandText = "SELECT  �λ���.�λ�,��Ա��.��ԱID, ��Ա��.����, ��Ա��.������λ, ��Ա��.��Ա���, A.��������, A.��λ��, A.����ID, ��Ա��.���, ��Ա��.ְ��, ��Ա��.�ֻ�, ��Ա��.EMail, A.����, A.�ʼ�, A.�ظ� FROM ��Ա�� LEFT OUTER JOIN (SELECT ��ԱID AS RID, ����ID, ��������, ��λ��, ����, �ʼ�, �ظ� FROM �λ���Ա�� WHERE (����ID = " + intConferenceID.ToString() + ")) A ON ��Ա��.��ԱID = A.RID CROSS JOIN �λ���";
                    sqlDA.Fill(dSet, "People");

                    sqlComm.CommandText = "SELECT ��ԱID FROM �λ���Ա�� WHERE (����ID = "+intConferenceID.ToString()+")";
                    sqlDA.Fill(dSet, "cPeople");//�λ���Ա�б�

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
                    dSet.Tables["People"].Columns.Remove("����ID");

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
            toolStripStatusLabelConference.Text = "�μӻ��������Ϊ��" + intCount.ToString() + "��";
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

            //�޸��б����
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
            //�޸��б�
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
                if (MessageBox.Show("�����Ѿ����棬�Ƿ������µĻ��飿", "�������", MessageBoxButtons.YesNo, MessageBoxIcon.Error) != DialogResult.Yes)
                    return;
            }


            System.Data.SqlClient.SqlDataReader sqldr;
            System.Data.SqlClient.SqlTransaction sqlta;
            int i;

            if (textBoxCname.Text.Trim() == "")
            {
                MessageBox.Show("�������Ʋ���Ϊ�գ�", "�������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.Connection = sqlConn;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;


            switch(intMode)
            {
                case 1://���ӻ���


                    try
                    {
                        sqlComm.CommandText = "INSERT INTO ����� (��������, ���鸱��, ��������, �᳡��, ��ʼʱ��, ����ʱ��) VALUES (N'" + textBoxCname.Text.Trim() + "', N'" + textBoxCname2.Text.Trim() + "', N'" + textBoxCintro.Text.Trim() + "', N'" + textBoxConferenceNo.Text.Trim() + "', '" + dateTimePickerStart.Value.ToString() + "', '" + dateTimePickerEnd .Value.ToString()+ "')";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "SELECT @@IDENTITY";
                        sqldr = sqlComm.ExecuteReader();
                        sqldr.Read();
                        string strCId = sqldr.GetValue(0).ToString();
                        intConferenceID = int.Parse(strCId);
                        sqldr.Close();


                        //�����
                        for (i = 0; i < listBoxGroup.Items.Count; i++)
                        {
                            sqlComm.CommandText = "INSERT INTO ����� (����ID, ��������) VALUES ("+strCId+", N'"+listBoxGroup.Items[i].ToString()+"')";
                            sqlComm.ExecuteNonQuery();
                        }

                        //�λ���Ա��
                        btnsAll_Click(null, null);
                        foreach(DataGridViewRow r in dataGridViewPeople.Rows)
                        {
                            if ((bool)r.Cells[0].Value)
                            {
                                sqlComm.CommandText = "INSERT INTO �λ���Ա�� (����ID, ��ԱID, ��������, ��λ��) VALUES (" + strCId + ", " + r.Cells[1].Value + ", N'" + r.Cells[5].Value + "', N'" + r.Cells[6].Value + "')";
                                sqlComm.ExecuteNonQuery();
                            }
                        }

                        sqlta.Commit();
                    }

                    catch (Exception err)
                    {
                        //�ع�
                        MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        sqlta.Rollback();
                        sqlConn.Close();
                        return;

                    }
                    finally
                    {
                        MessageBox.Show("�����Ѽ��뵽���ݿ⣡", "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    sqlConn.Close();

                    break;
                case 2://�޸Ļ���

                    try
                    {
                        sqlComm.CommandText = "UPDATE ����� SET �������� = N'" + textBoxCname.Text.Trim() + "', ���鸱�� = N'" + textBoxCname2.Text.Trim() + "', �������� = N'" + textBoxCintro.Text.Trim() + "', �᳡�� = N'" + textBoxConferenceNo.Text.Trim() + "', ��ʼʱ�� = '" + dateTimePickerStart.Value.ToString() + "', ����ʱ�� = '" + dateTimePickerEnd .Value.ToString()+ "' WHERE (����ID = "+intConferenceID.ToString()+")";
                        sqlComm.ExecuteNonQuery();



                        //�����
                        sqlComm.CommandText = "DELETE FROM ����� WHERE (����ID = " + intConferenceID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();
                        for (i = 0; i < listBoxGroup.Items.Count; i++)
                        {
                            sqlComm.CommandText = "INSERT INTO ����� (����ID, ��������) VALUES (" + intConferenceID.ToString() + ", N'" + listBoxGroup.Items[i].ToString() + "')";
                            sqlComm.ExecuteNonQuery();
                        }

                        //�λ���Ա��
                        System.Data.DataRow[] pRow;
                        //sqlComm.CommandText = "DELETE FROM �λ���Ա�� WHERE (����ID = " + intConferenceID.ToString() + ")";
                        //sqlComm.ExecuteNonQuery();
                        btnsAll_Click(null, null);
                        foreach (DataGridViewRow r in dataGridViewPeople.Rows)
                        {
                            pRow = dSet.Tables["cPeople"].Select("��ԱID = '" + r.Cells[1].Value.ToString() + "'");
                            if ((bool)r.Cells[0].Value) //ѡ�����Ա
                            {
                                if (pRow.Length < 1) //û�и���Ա����
                                {
                                    sqlComm.CommandText = "INSERT INTO �λ���Ա�� (����ID, ��ԱID, ��������, ��λ��) VALUES (" + intConferenceID.ToString() + ", " + r.Cells[1].Value + ", N'" + r.Cells[5].Value + "', N'" + r.Cells[6].Value + "')";
                                    
                                }
                                else
                                {
                                    sqlComm.CommandText = "UPDATE �λ���Ա�� SET �������� = N'" + r.Cells[5].Value + "', ��λ�� = N'" + r.Cells[6].Value + "' WHERE (��ԱID = " + r.Cells[1].Value + ") AND (����ID = "+intConferenceID.ToString()+")";
                                }
                                sqlComm.ExecuteNonQuery();
                            }
                            else //δѡ����Ա
                            {
                                if (pRow.Length > 0)
                                {
                                    sqlComm.CommandText = "DELETE FROM �λ���Ա�� WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID = " + r.Cells[1].Value.ToString() + ")";
                                    sqlComm.ExecuteNonQuery();
                                }
                                
                            }
                        }

                        sqlta.Commit();
                    }

                    catch (Exception err)
                    {
                        //�ع�
                        MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        sqlta.Rollback();
                        sqlConn.Close();
                        return;

                    }
                    finally
                    {
                        MessageBox.Show("������Ϣ���޸ģ�", "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                this.errorProviderConference.SetError(this.textBoxCname, "�ֶι���");
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
                this.errorProviderConference.SetError(this.textBoxCname2, "�ֶι���");
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
                this.errorProviderConference.SetError(this.textBoxConferenceNo, "�ֶι���");
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
                this.errorProviderConference.SetError(this.textBoxConferenceNo, "�ֶι���");
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
                this.errorProviderConference.SetError(this.textBoxConferenceNo, "�ֶι���");
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
                this.errorProviderConference.SetError(this.textBoxSite, "�ֶι���");
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
                this.errorProviderConference.SetError(this.textBoxGroup, "�ֶι���");
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
                sSearch = " ���� LIKE '%" + textBoxSXM.Text.Trim() + "%'";
                bFirst = false;
            }
            if (comboBoxSDW.Text.Trim() != "")
            {
                if (bFirst)
                    sSearch = " ������λ LIKE '%" + comboBoxSDW.Text.Trim() + "%'";
                else
                {
                    sSearch = sSearch + " AND ������λ LIKE '%" + comboBoxSDW.Text.Trim() + "%'";
                }
                bFirst = false;
            }
            if (textBoxXBH.Text.Trim() != "")
            {
                if (bFirst)
                    sSearch = " ��� LIKE '%" + textBoxXBH.Text.Trim() + "%'";
                else
                {
                    sSearch = sSearch + " AND ��� LIKE '%" + textBoxXBH.Text.Trim() + "%'";
                }
                bFirst = false;
            }
            if (comboBoxSZW.Text.Trim() != "")
            {
                if (bFirst)
                    sSearch = " ְ�� LIKE '%" + comboBoxSZW.Text.Trim() + "%'";
                else
                {
                    sSearch = sSearch + " AND ְ�� LIKE '%" + comboBoxSZW.Text.Trim() + "%'";
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
            dv1.RowFilter = " (�λ� = 1) ";

        }

        private void btnWCH_Click(object sender, EventArgs e)
        {

            DataView dv1 = new DataView(dSet.Tables["People"]);
            dataGridViewPeople.DataSource = dv1;
            dv1.RowStateFilter = DataViewRowState.CurrentRows;
            dv1.RowFilter = " (�λ� = 0) ";

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
                oledbDataAdapter.Fill(oledbdSet, "�����");
                oledbConn.Close();


                //��������
                System.Data.DataTable dtablePeople = oledbdSet.Tables["�����"];
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
                        //string sss1 = "����='" + dtablePeople.Rows[i][3].ToString() + "' AND ������λ='" + dtablePeople.Rows[i][2].ToString() + "'";

                        DataRow[] dr = dSet.Tables["People"].Select("����='" + dtablePeople.Rows[i][3].ToString() + "' AND ������λ='" + dtablePeople.Rows[i][2].ToString() + "'");

                        if (dr.Length > 0)
                        {
                            dr[0][0] = 1;
                        }

                    }
                }
                catch
                {
                    sqlConn.Close();
                    MessageBox.Show("���ݵ���ʧ�ܣ�", "���ݵ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                }
                MessageBox.Show("���ݵ���ɹ���", "���ݵ���", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                updateStateBar();

            }
            else
                return;
        }

        private void btnTZ_Click(object sender, EventArgs e)
        {
            if (intConferenceID == 0)
            {
                MessageBox.Show("���ȵ��ȷ����ť������飡", "����֪ͨ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                MessageBox.Show("����Ϣ:" + sms, "��ʾ��");
                sms = "";
            }
            if (id.Length != 0)
            {
                if (DialogResult.OK == MessageBox.Show(id + "����", "��ʾ��", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk))
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