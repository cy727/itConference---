using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.IO;


namespace itConference
{
    public partial class formCount : Form
    {

        public string strConn;
        public int intMode, intConferenceID;
        private System.Data.DataSet dSet = new DataSet();
        private System.Data.SqlClient.SqlDataReader sqldr;
        public string strVersion = "0";
        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;

        private int iLENGTH = 30;
        private int iCLENGTH = 10;

        private classCOM cCom = new classCOM();
        private ClassMember cMember = new ClassMember();

        public formCount()
        {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInfo.SelectedIndices.Count; i++)
            {
                j = listBoxInfo.SelectedIndices[i];
                listBoxInput.Items.Add(listBoxInfo.Items[j]);
                listBoxInfo.Items.RemoveAt(j);

            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInput.SelectedIndices.Count; i++)
            {
                j = listBoxInput.SelectedIndices[i];
                if (listBoxInput.Items[j].ToString() == "����") continue;
                listBoxInfo.Items.Add(listBoxInput.Items[j]);
                listBoxInput.Items.RemoveAt(j);

            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j < 1) return;
            string strSelect = listBoxInput.Items[j - 1].ToString();
            listBoxInput.Items[j - 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j - 1]);

        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j >= listBoxInput.Items.Count - 1) return;
            string strSelect = listBoxInput.Items[j + 1].ToString();
            listBoxInput.Items[j + 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j + 1]);

        }

        private void listBoxInfo_DoubleClick(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInfo.SelectedIndices.Count; i++)
            {
                j = listBoxInfo.SelectedIndices[i];
                listBoxInput.Items.Add(listBoxInfo.Items[j]);
                listBoxInfo.Items.RemoveAt(j);

            }
        }

        private void listBoxInput_DoubleClick(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInput.SelectedIndices.Count; i++)
            {
                j = listBoxInput.SelectedIndices[i];
                if (listBoxInput.Items[j].ToString() == "����") continue;
                listBoxInfo.Items.Add(listBoxInput.Items[j]);
                listBoxInput.Items.RemoveAt(j);

            }
        }

        private void btnOffInput_Click(object sender, EventArgs e)
        {
            string PID="";
            if (openFileDialogoff.ShowDialog() != DialogResult.OK) return;
            DateTime dtTemp, dtTemp1 = System.DateTime.Now, dtTemp2 = System.DateTime.Now;
            string stTemp1 = "", stTemp2 = "";
            
            
            //System.Data.SqlClient.SqlDataReader sqldr;

            if (dSet.Tables.Contains("���߻�����Ϣ")) dSet.Tables.Remove("���߻�����Ϣ");
            if (dSet.Tables.Contains("ǩ����")) dSet.Tables.Remove("ǩ����");

            try
            {
                XmlReadMode xmlM = dSet.ReadXml(openFileDialogoff.FileName, XmlReadMode.InferSchema);
            }
            catch
            {
                MessageBox.Show("��ȡǩ����Ϣ�ļ�����", "���ݿ�", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!dSet.Tables.Contains("ǩ����"))
            {
                MessageBox.Show("ǩ����Ϣ�ļ�����", "���ݿ�", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int i;

            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                for (i = 0; i < dSet.Tables["ǩ����"].Rows.Count; i++)
                {
                    if (dSet.Tables["ǩ����"].Rows[i][3].ToString() == "") //û��ǩ��ʱ��
                        continue;


                    //sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������,��Ա��.��ԱID, ��Ա��.��Ա���, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.����������, ��Ա��.��� FROM ��Ա�� WHERE (��Ա��.���� = N'" + dSet.Tables["ǩ����"].Rows[i][1].ToString() + "') AND (��Ա��.������λ = N'" + dSet.Tables["ǩ����"].Rows[i][2].ToString() + "')";
                    sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������,��Ա��.��ԱID, ��Ա��.��Ա���, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.����������, ��Ա��.��� FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0].ToString() + "') AND (������ <> N'')";
                    sqldr = sqlComm.ExecuteReader();

                    if (!sqldr.HasRows) //�޿���
                    {
                        sqldr.Close();
                        continue;
                    }

                    //��Ա��Ϣ
                    sqldr.Read();
                    PID = sqldr.GetValue(14).ToString();
                    sqldr.Close();

                    //sqlComm.CommandText = "SELECT �λ���Ա��.��ԱID FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')";
                    sqlComm.CommandText = "SELECT �λ���Ա��.��ԱID, �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.��ԱID = " + PID + ")";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows) //������Ա
                    {
                        sqldr.Read();

                        stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][3].ToString(), dSet.Tables["ǩ����"].Rows[i][4].ToString());
                        stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][3].ToString(), dSet.Tables["ǩ����"].Rows[i][4].ToString());

                        sqldr.Close();
                        if (stTemp2 != "")
                        {
                            //sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][7].ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = '" + stTemp2 + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                        }
                        else
                        {
                            //sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                        }
                        sqlComm.ExecuteNonQuery();
                    }
                    else //������Ա
                    {
                        sqldr.Close();
                        //sqlComm.CommandText = "SELECT �����µ���Ա��.��ԱID FROM �����µ���Ա�� INNER JOIN ��Ա�� ON �����µ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')";
                        sqlComm.CommandText = "SELECT �����µ���Ա��.��ԱID, �����µ���Ա��.ǩ��ʱ��, �����µ���Ա��.ǩ��ʱ�� FROM �����µ���Ա�� INNER JOIN ��Ա�� ON �����µ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.��ԱID = " + PID + ")";

                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows) //�м�¼
                        {
                            sqldr.Read();
                            stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][3].ToString(), dSet.Tables["ǩ����"].Rows[i][4].ToString());
                            stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][3].ToString(), dSet.Tables["ǩ����"].Rows[i][4].ToString());
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                //sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][7].ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = '" + stTemp2 + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                            }
                            else
                            {
                                //sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                            }
                            sqlComm.ExecuteNonQuery();
                        }
                        else //������¼
                        {
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                sqlComm.CommandText = "INSERT INTO �����µ���Ա�� (����ID, ��ԱID, ǩ��ʱ��, ǩ��ʱ��) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', '" + stTemp2 + "')";
                            }
                            else
                            {
                                sqlComm.CommandText = "INSERT INTO �����µ���Ա�� (����ID, ��ԱID, ǩ��ʱ��, ǩ��ʱ��) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', NULL)";
                            }
                            sqlComm.ExecuteNonQuery();
                        }

                    }
                }
                sqlta.Commit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("���ݿ����" + ex.Message.ToString(), "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                return;
            }
            finally
            {
                sqlConn.Close();
            }

            
            MessageBox.Show("ǩ����Ϣ������ϣ�", "���ݿ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }

        private void formCount_Load(object sender, EventArgs e)
        {
            int INFLEN,jj;

            INFLEN=30;
            jj = INFLEN >> 1;

            //�汾����
            switch (strVersion)
            {
                case "1": //������
                    MessageBox.Show("�ð汾��֧�ִ���ܣ��빺��߼��汾��", "�汾��Ϣ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "2": //��׼��
                case "3": //������
                    break;
                default:
                    break;
            } 
            
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlComm.CommandText = "SELECT ����ID, ��������, ���鸱��, ��ʼʱ��, ����ʱ��, �᳡�� FROM �����";
            if (intConferenceID != 0) sqlComm.CommandText = sqlComm.CommandText + " WHERE (����ID = " + intConferenceID.ToString() + ")";

            sqlConn.Open();
            sqlDA.Fill(dSet, "������Ϣ");

            sqlComm.CommandText = "SELECT ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            if (dSet.Tables.Contains("ϵͳ������")) dSet.Tables.Remove("ϵͳ������");
            sqlDA.Fill(dSet, "ϵͳ������");

            if (intConferenceID != 0) //��������ͳ��
            {
                if (dSet.Tables["������Ϣ"].Rows[0][2].ToString() != "")
                    labelConference.Text = dSet.Tables["������Ϣ"].Rows[0][1].ToString() + "��" + dSet.Tables["������Ϣ"].Rows[0][2].ToString();
                else
                    labelConference.Text = dSet.Tables["������Ϣ"].Rows[0][1].ToString();

                checkBoxhztjxx.Enabled = false;
                
            }
            else
            {
                labelConference.Text = "ȫ��������Ϣͳ��";
                btnCount.Enabled = false;
                btnOffInput.Enabled = false;
                btnSCJ.Enabled = false;
            }
            sqlConn.Close();

            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCount_Click(object sender, EventArgs e)
        {
            string sTemp = "";

            //ȫ������
            if (intConferenceID == 0) return;

            sTemp = "SELECT ��Ա��.��ԱID, ";
            int i;
            for (i = 0; i < listBoxInput.Items.Count; i++)
            {
                sTemp = sTemp + "��Ա��." + listBoxInput.Items[i].ToString() + ", ";
            }


            //ǩ��
            sqlComm.CommandText = sTemp + " �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (�λ���Ա��.ǩ��ʱ�� IS NOT NULL) ORDER BY ǩ��ʱ�� DESC";


            if (dSet.Tables.Contains("SIGN"))
            {
                dSet.Tables.Remove("SIGN");
                dataGridViewSign.DataSource = "";
            }
            sqlDA.Fill(dSet, "SIGN");

            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            dataGridViewSign.Columns[0].Visible = false;
            for (i = 1; i < listBoxInput.Items.Count; i++)
            {
                dataGridViewSign.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            dataGridViewSign.Columns[i+1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[i + 2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //ȱϯ
            sqlComm.CommandText = sTemp + " �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (�λ���Ա��.ǩ��ʱ�� IS NULL) ORDER BY ǩ��ʱ�� DESC";


            if (dSet.Tables.Contains("UNSIGN"))
            {
                dSet.Tables.Remove("UNSIGN");
                dataGridViewUnSign.DataSource = "";
            }
            sqlDA.Fill(dSet, "UNSIGN");

            dataGridViewUnSign.DataSource = dSet.Tables["UNSIGN"];
            dataGridViewUnSign.Columns[0].Visible = false;
            for (i = 1; i < listBoxInput.Items.Count; i++)
            {
                dataGridViewUnSign.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            dataGridViewUnSign.Columns[i + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewUnSign.Columns[i + 2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewUnSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //����
            sqlComm.CommandText = sTemp + " �����µ���Ա��.ǩ��ʱ��, �����µ���Ա��.ǩ��ʱ�� FROM �����µ���Ա�� INNER JOIN ��Ա�� ON �����µ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (�����µ���Ա��.ǩ��ʱ�� IS NOT NULL) ORDER BY ǩ��ʱ�� DESC";


            if (dSet.Tables.Contains("NEW"))
            {
                dSet.Tables.Remove("NEW");
                dataGridViewNew.DataSource = "";
            }
            sqlDA.Fill(dSet, "NEW");

            dataGridViewNew.DataSource = dSet.Tables["NEW"];
            dataGridViewNew.Columns[0].Visible = false;
            for (i = 1; i < listBoxInput.Items.Count; i++)
            {
                dataGridViewNew.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            dataGridViewNew.Columns[i + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewNew.Columns[i + 2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewNew.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            changestatusBar();


        }

        private void changestatusBar()
        {

            System.Data.SqlClient.SqlDataReader sqldr;
            string sText="";
            int iTemp=0;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT COUNT(*) FROM �λ���Ա�� GROUP BY ����ID HAVING (����ID = " + dSet.Tables["������Ϣ"].Rows[0][0] + ")";
            sqldr = sqlComm.ExecuteReader();
            if(sqldr.HasRows)
            {
                sqldr.Read();
                sText = "��¼������ "+sqldr.GetValue(0).ToString()+" ��";
                iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
             }
            sqldr.Close();

            
             sqlComm.CommandText = "SELECT COUNT(*) FROM �λ���Ա�� WHERE (����ID = " + dSet.Tables["������Ϣ"].Rows[0][0] +") AND (ǩ��ʱ�� IS NOT NULL)";
             sqldr = sqlComm.ExecuteReader();
            if(sqldr.HasRows)
            {
                sqldr.Read();
                sText = sText+",ǩ�������� "+sqldr.GetValue(0).ToString()+" ��,ȱϯ������ " + (iTemp - Int32.Parse(sqldr.GetValue(0).ToString())).ToString() + " ��,����������"+dataGridViewNew.Rows.Count.ToString()+"��";;
                sqldr.Close();
             }

            sqlConn.Close();
            toolStripStatusLabelCount.Text = sText;
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlDataReader sqldr;
            
            if (saveFileDialogOutput.ShowDialog() != DialogResult.OK) return;


            //����һ��ExcelӦ�ó���ʵ��
            sqlConn.Open();
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            if (objExcel == null) 
            {
                sqlConn.Close();
                MessageBox.Show("�޷�����Excel�ĵ���", "��������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            objExcel.Visible = false; 

            //����һ��Excel�ļ���δ���棬���ļ�����
            Workbooks objWorkbooks = objExcel.Workbooks;
            _Workbook objWorkbook = objWorkbooks.Add(XlWBATemplate.xlWBATWorksheet);//Ĭ�ϴ���sheet1

            int i, j, k, iTemp=0,intCount=0;
            Microsoft.Office.Interop.Excel.Range range ;

            for (i = 0; i < dSet.Tables["������Ϣ"].Rows.Count; i++)
            {
                if(i>0) objWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, XlWBATemplate.xlWBATWorksheet);
                Sheets objSheets = objWorkbook.Worksheets;

                //�������
                //_Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(i+1);
                _Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(1);
                objWorksheet.Name = dSet.Tables["������Ϣ"].Rows[i][1].ToString();

                range = objWorksheet.get_Range("A1", Missing.Value);
                range[1, 1] = "ǩ���������";

                if (dSet.Tables["������Ϣ"].Rows[i][2].ToString()=="") //���鸱��Ϊ��
                    range[3, 1] = "�������ƣ� " + dSet.Tables["������Ϣ"].Rows[i][1];
                else
                    range[3, 1] = "�������ƣ� " + dSet.Tables["������Ϣ"].Rows[i][1] + "��" + dSet.Tables["������Ϣ"].Rows[i][2];
                range[4, 1] = "��ʼʱ�䣺 " + dSet.Tables["������Ϣ"].Rows[i][3] + " ����ʱ�䣺 " + dSet.Tables["������Ϣ"].Rows[i][4];

                intCount = 5; //����м���
                sqlComm.CommandText = "SELECT COUNT(*) FROM �λ���Ա�� GROUP BY ����ID HAVING (����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ")";
                sqldr = sqlComm.ExecuteReader();

                if(sqldr.HasRows)
                {
                    sqldr.Read();
                    if (checkBoxderysl.Checked)
                    {
                        range[intCount, 1] = "��¼������ " + sqldr.GetValue(0).ToString() + " ��";
                        intCount++;
                    }
                    iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
                }
                sqldr.Close();


                sqlComm.CommandText = "SELECT COUNT(*) FROM �λ���Ա�� WHERE (����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (ǩ��ʱ�� IS NOT NULL)";
               // sqlComm.CommandText = "SELECT COUNT(*) FROM �λ���Ա�� GROUP BY ����ID, ǩ��ʱ�� HAVING (����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (ǩ��ʱ�� IS NOT NULL)";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    if (checkBoxqdrysl.Checked)
                    {
                        range[intCount, 1] = "ǩ�������� " + sqldr.GetValue(0).ToString().ToString() + " ��";
                        intCount++;
                    }
                    if (checkBoxqxrysl.Checked)
                    {
                        range[intCount, 1] = "ȱϯ������ " + (iTemp - Int32.Parse(sqldr.GetValue(0).ToString())).ToString() + " ��";
                        intCount++;
                    }
                    iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
                }
                sqldr.Close();

                if (checkBoxNew.Checked)
                {

                    sqlComm.CommandText = "SELECT COUNT(*) FROM �����µ���Ա�� WHERE (����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (ǩ��ʱ�� IS NOT NULL)";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        range[intCount, 1] = "���������� " + sqldr.GetValue(0).ToString().ToString() + " ��";
                        intCount++;
                    }
                    sqldr.Close();
                }


                sqlComm.CommandText = "SELECT COUNT(*) FROM �λ���Ա�� WHERE (����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (ǩ��ʱ�� > CONVERT(DATETIME, '" + dSet.Tables["������Ϣ"].Rows[i][3] + "', 102))";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    if (checkBoxcdrysl.Checked)
                    {
                        iTemp = iTemp - Int32.Parse(sqldr.GetValue(0).ToString());
                        range[intCount, 1] = "׼ʱ��� " + iTemp.ToString() + " �ˣ� �ٵ��� " + sqldr.GetValue(0).ToString() + " ��";
                        intCount++;
                    }
                }
                sqldr.Close();

                //ǩ��
                intCount++;

                if(checkBoxcdrylb.Checked) //��ʾ�ٵ���Ա�б�
                {
                if (dSet.Tables.Contains("ǩ����Ϣ")) dSet.Tables.Remove("ǩ����Ϣ");
                sqlComm.CommandText = "SELECT ��Ա��.��ԱID, ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + "��Ա��." + listBoxInput.Items[j].ToString() + ", ";
                }
                if(checkBoxqcsj.Checked) //ǩ��ʱ��ͳ��
                    sqlComm.CommandText = sqlComm.CommandText + " �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�λ���Ա��.ǩ��ʱ�� <= CONVERT(DATETIME,  '" + dSet.Tables["������Ϣ"].Rows[i][3] + "', 102)) ORDER BY �λ���Ա��.ǩ��ʱ��";
                else
                    sqlComm.CommandText = sqlComm.CommandText + " �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�λ���Ա��.ǩ��ʱ�� <= CONVERT(DATETIME,  '" + dSet.Tables["������Ϣ"].Rows[i][3] + "', 102)) ORDER BY �λ���Ա��.ǩ��ʱ��";

                sqlDA.Fill(dSet, "ǩ����Ϣ");
                if (dSet.Tables["ǩ����Ϣ"].Rows.Count > 0)
                {
                    range[intCount, 1] = "��Աǩ�������";
                    range[intCount + 1, 1] = "���";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "��������һ":
                                if(dSet.Tables["ϵͳ������"].Rows.Count<1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][0].ToString();
                                break;

                            case "�������Զ�":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][1].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][2].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }

                        
                    }
                    range[intCount + 1, listBoxInput.Items.Count + 2] = "ǩ��ʱ��";
                    if(checkBoxqcsj.Checked) range[intCount + 1, listBoxInput.Items.Count + 3] = "ǩ��ʱ��";

                    for (j = 1; j < dSet.Tables["ǩ����Ϣ"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["ǩ����Ϣ"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["ǩ����Ϣ"].Rows[j - 1][k].ToString();
                        }
                    }
                    intCount += dSet.Tables["ǩ����Ϣ"].Rows.Count + 3;

                }


                //�ٵ�

                if (dSet.Tables.Contains("ǩ����Ϣ")) dSet.Tables.Remove("ǩ����Ϣ");
                sqlComm.CommandText = "SELECT ��Ա��.��ԱID, ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + "��Ա��." + listBoxInput.Items[j].ToString() + ", ";
                }
                sqlComm.CommandText = sqlComm.CommandText + " �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�λ���Ա��.ǩ��ʱ�� > CONVERT(DATETIME,  '" + dSet.Tables["������Ϣ"].Rows[i][3] + "', 102)) ORDER BY �λ���Ա��.ǩ��ʱ��";
                sqlDA.Fill(dSet, "ǩ����Ϣ");
                if (dSet.Tables["ǩ����Ϣ"].Rows.Count > 0)
                {
                    range[intCount, 1] = "�ٵ���Աǩ�������";
                    range[intCount + 1, 1] = "���";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "��������һ":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][0].ToString();
                                break;

                            case "�������Զ�":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][1].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][2].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }
                        
                        //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                    }
                    range[intCount + 1, listBoxInput.Items.Count + 2] = "ǩ��ʱ��";
                    range[intCount + 1, listBoxInput.Items.Count + 3] = "ǩ��ʱ��";

                    for (j = 1; j < dSet.Tables["ǩ����Ϣ"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["ǩ����Ϣ"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["ǩ����Ϣ"].Rows[j - 1][k].ToString();
                        }
                    }
                    intCount += dSet.Tables["ǩ����Ϣ"].Rows.Count + 3;

                }
                }
                else //����ʾ�ٵ��б�
                {
                if (dSet.Tables.Contains("ǩ����Ϣ")) dSet.Tables.Remove("ǩ����Ϣ");
                sqlComm.CommandText = "SELECT ��Ա��.��ԱID, ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + "��Ա��." + listBoxInput.Items[j].ToString() + ", ";
                }
                if(checkBoxqcsj.Checked) //ǩ��ʱ��ͳ��
                    sqlComm.CommandText = sqlComm.CommandText + " �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�λ���Ա��.ǩ��ʱ�� IS NOT NULL) ORDER BY �λ���Ա��.ǩ��ʱ��";
                else
                    sqlComm.CommandText = sqlComm.CommandText + " �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�λ���Ա��.ǩ��ʱ�� IS NOT NULL) ORDER BY �λ���Ա��.ǩ��ʱ��";

                sqlDA.Fill(dSet, "ǩ����Ϣ");
                if (dSet.Tables["ǩ����Ϣ"].Rows.Count > 0)
                {
                    range[intCount, 1] = "��Աǩ�������";
                    range[intCount + 1, 1] = "���";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "��������һ":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][0].ToString();
                                break;

                            case "�������Զ�":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][1].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][2].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }

                        //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                    }
                    range[intCount + 1, listBoxInput.Items.Count + 2] = "ǩ��ʱ��";
                    if(checkBoxqcsj.Checked) range[intCount + 1, listBoxInput.Items.Count + 3] = "ǩ��ʱ��";

                    for (j = 1; j < dSet.Tables["ǩ����Ϣ"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["ǩ����Ϣ"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["ǩ����Ϣ"].Rows[j - 1][k].ToString();
                        }
                    }
                    intCount += dSet.Tables["ǩ����Ϣ"].Rows.Count + 3;

                }

                }

                //ȱϯ
                if(checkBoxqxrylb.Checked)
                {
                if (dSet.Tables.Contains("ǩ����Ϣ")) dSet.Tables.Remove("ǩ����Ϣ");
                sqlComm.CommandText = "SELECT ��Ա��.��ԱID ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + ", ��Ա��." + listBoxInput.Items[j].ToString() ;
                }
                sqlComm.CommandText = sqlComm.CommandText + " FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�λ���Ա��.ǩ��ʱ�� IS NULL)";
                sqlDA.Fill(dSet, "ǩ����Ϣ");
                if (dSet.Tables["ǩ����Ϣ"].Rows.Count > 0)
                {
                    range[intCount, 1] = "ȱϯ��Ա��";
                    range[intCount + 1, 1] = "���";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "��������һ":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][0].ToString();
                                break;

                            case "�������Զ�":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][1].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][2].ToString();
                                break;

                            case "����������":
                                if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }
                        
                        //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                    }


                    for (j = 1; j < dSet.Tables["ǩ����Ϣ"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["ǩ����Ϣ"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["ǩ����Ϣ"].Rows[j - 1][k].ToString();
                        }
                    }
                 }
                }


                //����
                if (checkBoxNew.Checked)
                {

                    if (dSet.Tables.Contains("������Ϣ")) dSet.Tables.Remove("������Ϣ");
                    sqlComm.CommandText = "SELECT ��Ա��.��ԱID ";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        sqlComm.CommandText = sqlComm.CommandText + ", ��Ա��." + listBoxInput.Items[j].ToString();
                    }
                    sqlComm.CommandText = sqlComm.CommandText + ", �����µ���Ա��.ǩ��ʱ��, �����µ���Ա��.ǩ��ʱ��  FROM �����µ���Ա�� INNER JOIN ��Ա�� ON �����µ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + dSet.Tables["������Ϣ"].Rows[i][0] + ") AND (�����µ���Ա��.ǩ��ʱ�� IS NOT NULL)";
                    sqlDA.Fill(dSet, "������Ϣ");
                    if (dSet.Tables["������Ϣ"].Rows.Count > 0)
                    {
                        range[intCount, 1] = "������Ա��";
                        range[intCount + 1, 1] = "���";
                        for (j = 0; j < listBoxInput.Items.Count; j++)
                        {
                            switch (listBoxInput.Items[j].ToString())
                            {
                                case "��������һ":
                                    if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][0].ToString();
                                    break;

                                case "�������Զ�":
                                    if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][1].ToString();
                                    break;

                                case "����������":
                                    if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][2].ToString();
                                    break;

                                case "����������":
                                    if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["ϵͳ������"].Rows[0][3].ToString();
                                    break;

                                default:
                                    range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                    break;
                            }

                            //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                        }
                        range[intCount + 1, listBoxInput.Items.Count + 2] = "ǩ��ʱ��";
                        if (checkBoxqcsj.Checked) range[intCount + 1, listBoxInput.Items.Count + 3] = "ǩ��ʱ��";


                        for (j = 1; j < dSet.Tables["������Ϣ"].Rows.Count + 1; j++)
                        {
                            range[j + intCount + 1, 1] = j.ToString();
                            for (k = 1; k < dSet.Tables["������Ϣ"].Columns.Count; k++)
                            {
                                if (!checkBoxqcsj.Checked && k == dSet.Tables["������Ϣ"].Columns.Count - 1)
                                    continue;
                                range[j + intCount + 1, k + 1] = dSet.Tables["������Ϣ"].Rows[j - 1][k].ToString();
                            }
                        }
                    }
                }




            }

            if(intConferenceID==0) //ȫ������ͳ��
            {
                if (checkBoxhztjxx.Checked) //����ͳ����Ϣ
                {
                    objWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, XlWBATemplate.xlWBATWorksheet);
                    Sheets objSheets = objWorkbook.Worksheets;

                    //�������
                    _Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(1);
                    objWorksheet.Name = "ȫ���������ͳ�Ʊ���";

                    range = objWorksheet.get_Range("A1", Missing.Value);
                    sqlComm.CommandText = "SELECT COUNT(*) AS [COUNT], �λ���Ա��.����ID, �����.��������, �����.���鸱�� FROM �λ���Ա�� INNER JOIN  ����� ON �λ���Ա��.����ID = �����.����ID WHERE (�λ���Ա��.ǩ��ʱ�� IS NOT NULL) GROUP BY �λ���Ա��.����ID, �����.��������, �����.���鸱�� ORDER BY COUNT(*) DESC";
                    if (dSet.Tables.Contains("����ͳ��")) dSet.Tables.Remove("����ͳ��");
                    sqlDA.Fill(dSet, "����ͳ��");
                    intCount = 1;
                    if (dSet.Tables["����ͳ��"].Rows.Count > 0)
                    {
                        range[intCount, 1] = "ȫ������ǩ��ͳ�ƣ�";
                        range[intCount + 1, 1] = "���";
                        range[intCount + 1, 2] = "��������";
                        range[intCount + 1, 3] = "���鸱��";
                        range[intCount + 1, 4] = "ǩ������";
                        if (checkBoxNew.Checked)
                        {
                            range[intCount + 1, 5] = "��������";
                        }

                        for (j = 1; j < dSet.Tables["����ͳ��"].Rows.Count + 1; j++)
                        {
                            range[j + intCount + 1, 1] = j.ToString();
                            range[j + intCount + 1, 2] = dSet.Tables["����ͳ��"].Rows[j - 1][2].ToString();
                            range[j + intCount + 1, 3] = dSet.Tables["����ͳ��"].Rows[j - 1][3].ToString();
                            range[j + intCount + 1, 4] = dSet.Tables["����ͳ��"].Rows[j - 1][0].ToString();

                            if (checkBoxNew.Checked)
                            {

                                sqlComm.CommandText = "SELECT COUNT(*) FROM �����µ���Ա�� WHERE (����ID = " + dSet.Tables["����ͳ��"].Rows[j - 1][1].ToString() + ")";
                                sqldr = sqlComm.ExecuteReader();
                                if (sqldr.HasRows)
                                {
                                    sqldr.Read();
                                    range[j + intCount + 1, 5] = sqldr.GetValue(0).ToString();

                                }
                                else
                                {
                                    range[j + intCount + 1, 5] = "0";
                                }
                                sqldr.Close();
                            }

                        }
                    }
                    
                   
                }
            }

            //д������
 

            //�����ļ�
            objWorkbook._SaveAs(saveFileDialogOutput.FileName, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            //�ر��ļ�
            objWorkbook.Close(false, saveFileDialogOutput.FileName, false);
            //objExcel.Visible = false;
            objExcel.Quit();
            objExcel = null;

            sqlConn.Close();
            MessageBox.Show("�ɹ�����Excel�ĵ���" + saveFileDialogOutput.FileName.ToString(), "����", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

        }

        private void btnSCJ_Click(object sender, EventArgs e)
        {
            formSCJ frmSCJ=new formSCJ();
            if (frmSCJ.ShowDialog() != DialogResult.OK)
                return;

            int i, iR;
            string sTemp = "";

            //ɾ����ʱ�ļ�

            string dFileName = Directory.GetCurrentDirectory() + "\\ittmp.txt";
            if (File.Exists(dFileName))
            {
                File.Delete(dFileName);
            }

            cCom.sComPort = frmSCJ.comboBoxCOM.Text;
            iR = cCom.UpFile(Directory.GetCurrentDirectory());

            if (iR != 0)
            {
                MessageBox.Show("���ֳֻ���������ʧ��");
                return;
            }
 

            string PID = "", sTemp1="";
            DateTime dtTemp, dtTemp1 = System.DateTime.Now, dtTemp2 = System.DateTime.Now;
            string stTemp1 = "", stTemp2 = "";
            DataRow[] dtC;
            bool bDT;
 
            if (dSet.Tables.Contains("���߻�����Ϣ")) dSet.Tables.Remove("���߻�����Ϣ");
            if (dSet.Tables.Contains("ǩ����")) dSet.Tables.Remove("ǩ����");

            dSet.Tables.Add("ǩ����");

            dSet.Tables["ǩ����"].Columns.Add("��Ա����", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("ǩ��ʱ��", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("ǩ��ʱ��", System.Type.GetType("System.String"));

            StreamReader swTemp = new StreamReader(dFileName, Encoding.Default);
            while (swTemp.Peek() >= 0)
            {
                string[] strDRow = {"", "", ""};
                sTemp = swTemp.ReadLine();
                if (sTemp.Length < iLENGTH)
                    continue;

                stTemp1 = sTemp.Substring(0, iCLENGTH).Trim();

                if (intCard == 0) //ID��������ת��
                    strDRow[0] = cMember.IDCARDNOFROMHEX(stTemp1, intCardLength);
                else
                    strDRow[0] = stTemp1;

                if (strDRow[0] == "0000000000")
                    continue;

                sTemp1 = sTemp.Substring(iCLENGTH + 1, 19);
                try
                {
                    dtTemp = DateTime.Parse(sTemp1);
                }
                catch
                {
                    continue;
                }

                //����Ƿ�����
                dtC = dSet.Tables["ǩ����"].Select("��Ա���� = '" + strDRow[0] + "'");
                if (dtC.Length > 0) //����ǩ��
                {
                    //ǩ��ʱ��
                    try
                    {
                        dtTemp1 = DateTime.Parse(dtC[0][1].ToString());
                    }
                    catch
                    {
                        continue;
                    }

                    //ǩ��ʱ��
                    bDT = true;
                    try
                    {
                        dtTemp2 = DateTime.Parse(dtC[0][2].ToString());
                    }
                    catch
                    {
                        bDT=false;
                    }

                    if (bDT) //��ǩ��ʱ��
                    {
                        if (dtTemp < dtTemp1)
                        {
                            dtC[0][1] = dtTemp.ToString();
                        }
                        else
                        {
                            if (dtTemp > dtTemp2)
                            dtC[0][2] = dtTemp.ToString();
                        }


                    }
                    else //û��ǩ��ʱ��
                    {
                        if(dtTemp<dtTemp1)
                        {
                            dtC[0][1]=dtTemp.ToString();
                            dtC[0][2]=dtTemp1.ToString();
                        }
                        else
                        {
                            dtC[0][2]=dtTemp.ToString();
                        }
                    }
                }
                else
                {
                    strDRow[1] = sTemp1;
                    dSet.Tables["ǩ����"].Rows.Add(strDRow);
                }
                
            }


            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {

                for (i = 0; i < dSet.Tables["ǩ����"].Rows.Count; i++)
                {
                    if (dSet.Tables["ǩ����"].Rows[i][1].ToString() == "") //û��ǩ��ʱ��
                        continue;


                    //sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������,��Ա��.��ԱID, ��Ա��.��Ա���, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.����������, ��Ա��.��� FROM ��Ա�� WHERE (��Ա��.���� = N'" + dSet.Tables["ǩ����"].Rows[i][1].ToString() + "') AND (��Ա��.������λ = N'" + dSet.Tables["ǩ����"].Rows[i][2].ToString() + "')";
                    sqlComm.CommandText = "SELECT ��Ա��.����, ��Ա��.�Ա�, ��Ա��.����, ��Ա��.ְ��, ��Ա��.ְ��, ��Ա��.������λ, ��Ա��.����, ��Ա��.�ֻ�, ��Ա��.�绰, ��Ա��.����, ��Ա��.EMail, ��Ա��.ͨѶ��ַ, ��Ա��.��������, ��Ա��.������,��Ա��.��ԱID, ��Ա��.��Ա���, ��Ա��.��������һ, ��Ա��.�������Զ�, ��Ա��.����������, ��Ա��.����������, ��Ա��.��� FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0].ToString() + "') AND (������ <> N'')";
                    sqldr = sqlComm.ExecuteReader();

                    if (!sqldr.HasRows) //�޿���
                    {
                        sqldr.Close();
                        continue;
                    }

                    //��Ա��Ϣ
                    sqldr.Read();
                    PID = sqldr.GetValue(14).ToString();
                    sqldr.Close();

                    //sqlComm.CommandText = "SELECT �λ���Ա��.��ԱID FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')";
                    sqlComm.CommandText = "SELECT �λ���Ա��.��ԱID, �λ���Ա��.ǩ��ʱ��, �λ���Ա��.ǩ��ʱ�� FROM �λ���Ա�� INNER JOIN ��Ա�� ON �λ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�λ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.��ԱID = " + PID + ")";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows) //������Ա
                    {
                        sqldr.Read();

                        stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][1].ToString(), dSet.Tables["ǩ����"].Rows[i][2].ToString());
                        stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][1].ToString(), dSet.Tables["ǩ����"].Rows[i][2].ToString());

                        sqldr.Close();
                        if (stTemp2 != "")
                        {
                            //sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][7].ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = '" + stTemp2 + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                        }
                        else
                        {
                            //sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE �λ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                        }
                        sqlComm.ExecuteNonQuery();
                    }
                    else //������Ա
                    {
                        sqldr.Close();
                        //sqlComm.CommandText = "SELECT �����µ���Ա��.��ԱID FROM �����µ���Ա�� INNER JOIN ��Ա�� ON �����µ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')";
                        sqlComm.CommandText = "SELECT �����µ���Ա��.��ԱID, �����µ���Ա��.ǩ��ʱ��, �����µ���Ա��.ǩ��ʱ�� FROM �����µ���Ա�� INNER JOIN ��Ա�� ON �����µ���Ա��.��ԱID = ��Ա��.��ԱID WHERE (�����µ���Ա��.����ID = " + intConferenceID.ToString() + ") AND (��Ա��.��ԱID = " + PID + ")";

                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows) //�м�¼
                        {
                            sqldr.Read();

                            stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][1].ToString(), dSet.Tables["ǩ����"].Rows[i][2].ToString());
                            stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["ǩ����"].Rows[i][1].ToString(), dSet.Tables["ǩ����"].Rows[i][2].ToString());
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                //sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][7].ToString() + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = '" + stTemp2 + "' WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                            }
                            else
                            {
                                //sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + dSet.Tables["ǩ����"].Rows[i][6].ToString() + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (������ = N'" + dSet.Tables["ǩ����"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE �����µ���Ա�� SET ǩ��ʱ�� = '" + stTemp1 + "', ǩ��ʱ�� = NULL WHERE (����ID = " + intConferenceID.ToString() + ") AND (��ԱID IN (SELECT ��ԱID FROM ��Ա�� WHERE (��ԱID = " + PID + ")))";
                            }
                            sqlComm.ExecuteNonQuery();
                        }
                        else //������¼
                        {
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                sqlComm.CommandText = "INSERT INTO �����µ���Ա�� (����ID, ��ԱID, ǩ��ʱ��, ǩ��ʱ��) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', '" + stTemp2 + "')";
                            }
                            else
                            {
                                sqlComm.CommandText = "INSERT INTO �����µ���Ա�� (����ID, ��ԱID, ǩ��ʱ��, ǩ��ʱ��) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', NULL)";
                            }
                            sqlComm.ExecuteNonQuery();
                        }

                    }
                }
                sqlta.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("���ݿ����" + ex.Message.ToString(), "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                return;
            }
            finally
            {
                sqlConn.Close();
            }


            MessageBox.Show("ǩ����Ϣ������ϣ�", "���ݿ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private string strMin(string s1, string s2, string s3, string s4)
        {
            if (s1.Trim() == "" && s2.Trim() == "" && s3.Trim() == "" && s4.Trim() == "")
                return "";

            string sTemp;
            DateTime dTemp1,dTemp2,dTemp;

            DateTime dNULL=DateTime.Parse("1900-1-1");
            dTemp1=dNULL;
            dTemp2=dNULL;
            dTemp=dNULL;


            try
            {
                dTemp1 = DateTime.Parse(s1);
            }
            catch
            {
                dTemp1 = dNULL;
            }

            try
            {
                dTemp = DateTime.Parse(s2);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp < dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s3);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp < dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s4);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp < dTemp1)
                        dTemp1 = dTemp;
                }
            }

            if (dTemp1 == dNULL)
                return "";
            else
                return dTemp1.ToString();

        }


        private string strMax(string s1, string s2, string s3, string s4)
        {
            int i;
            i = 0;
            if (s1.Trim() == "")
                i++;
            if (s2.Trim() == "")
                i++;
            if (s3.Trim() == "")
                i++;
            if (s4.Trim() == "")
                i++;

            if (i >= 3)
                return "";


            string sTemp;
            DateTime dTemp1, dTemp2, dTemp;
            DateTime dNULL = DateTime.Parse("1900-1-1");
            dTemp1 = dNULL;
            dTemp2 = dNULL;
            dTemp = dNULL;
            

            try
            {
                dTemp1 = DateTime.Parse(s1);
            }
            catch
            {
                dTemp1 = dNULL;
            }

            try
            {
                dTemp = DateTime.Parse(s2);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp > dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s3);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp > dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s4);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp > dTemp1)
                        dTemp1 = dTemp;
                }
            }

            if (dTemp1 == dNULL)
                return "";
            else
                return dTemp1.ToString();
        }
    }
}