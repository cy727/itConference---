using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;


namespace itConference
{
    public partial class mainCon : Form
    {
        public string strConn="";
        private int intMenu , intContent;
        private int intComNumber = 1, intCardLength=10, intCard=1;
        private int lComBand = 19200;
        private System.Data.DataSet dSet=new DataSet();
        private string strPSelect;
        private string strSerial="";
        private string strVersion = "0";

        private classCOM cCom = new classCOM();
        private ClassMember cMember = new ClassMember();

        public mainCon()
        {
            InitializeComponent();
            
            //��ʼ��

            tableLayoutP1.Visible = false;
            tableLayoutP2.Visible = false;
            tableLayoutP3.Visible = false;
            tableLayoutP4.Visible = false;
            tableLayoutP5.Visible = false;
            tableLayoutP6.Visible = false;
            tableLayoutP1.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutP2.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutP3.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutP4.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutP5.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutP6.Dock = System.Windows.Forms.DockStyle.Fill;

            intMenu = 0;intContent = 0;
            strPSelect = "SELECT ��ԱID, ����, �Ա�, ������λ, ְ��, ְ��, ͨѶ��ַ, ��Ա���, ������ FROM ��Ա��";

            

        }

        private void initInteface()
        {
            int i;

            if (intMenu <  listViewMain.Items.Count-1) //���˳�
                toolStripStatusLabelS.Text = listViewMain.Items[intMenu].Text;

            if (intMenu != 0) //ɾ����Ա�˵�
                DelPeopleToolStripMenuItem.Enabled = false;
            else
                DelPeopleToolStripMenuItem.Enabled = true;
            

            if (intMenu != 1) //ɾ������˵�
                DelConferenceToolStripMenuItem.Enabled = false;
            else
                DelConferenceToolStripMenuItem.Enabled = true;

            switch(intMenu)
            {
                case 0:
                    tableLayoutP1.Visible = true;
                    tableLayoutP2.Visible = false;
                    tableLayoutP3.Visible = false;
                    tableLayoutP4.Visible = false;
                    tableLayoutP5.Visible = false;
                    tableLayoutP6.Visible = false;
                    break;
                case 1:
                    tableLayoutP1.Visible = false;
                    tableLayoutP2.Visible = true;
                    tableLayoutP3.Visible = false;
                    tableLayoutP4.Visible = false;
                    tableLayoutP5.Visible = false;
                    tableLayoutP6.Visible = false; 
                    break;
                case 2:
                    tableLayoutP1.Visible = false;
                    tableLayoutP2.Visible = false;
                    tableLayoutP3.Visible = true;
                    tableLayoutP4.Visible = false;
                    tableLayoutP5.Visible = false;
                    tableLayoutP6.Visible = false;
                    break;
                case 3:
                    tableLayoutP1.Visible = false;
                    tableLayoutP2.Visible = false;
                    tableLayoutP3.Visible = false;
                    tableLayoutP4.Visible = true;
                    tableLayoutP5.Visible = false;
                    tableLayoutP6.Visible = false;
                    break;
                case 4:
                    tableLayoutP1.Visible = false;
                    tableLayoutP2.Visible = false;
                    tableLayoutP3.Visible = false;
                    tableLayoutP4.Visible = false;
                    tableLayoutP5.Visible = true;
                    tableLayoutP6.Visible = false;
                    break;
                case 5:
                    tableLayoutP1.Visible = false;
                    tableLayoutP2.Visible = false;
                    tableLayoutP3.Visible = false;
                    tableLayoutP4.Visible = false;
                    tableLayoutP5.Visible = false;
                    tableLayoutP6.Visible = true;
                    break;
                default:
                    tableLayoutP1.Visible = false;
                    tableLayoutP2.Visible = false;
                    tableLayoutP3.Visible = false;
                    tableLayoutP4.Visible = false;
                    tableLayoutP5.Visible = false;
                    tableLayoutP6.Visible = false;
                    break;
            }

            for (i = 0; i < listViewMain.Items.Count; i++)
            {
                listViewMain.Items[i].Selected = false;
                listViewMain.Items[i].ImageIndex = i;
            }
            listViewMain.Items[intMenu].ImageIndex = intMenu + listViewMain.Items.Count;


            if (dataGridViewPeople.Columns.Count > 0)
                dataGridViewPeople.Columns[0].Visible = false;
            if (dataGridViewConference.Columns.Count > 0)
                dataGridViewConference.Columns[0].Visible = false;
            if (dataGridViewSign.Columns.Count > 0)
                dataGridViewSign.Columns[0].Visible = false;
            if (dataGridViewConsume.Columns.Count > 0)
                dataGridViewConsume.Columns[0].Visible = false;
            if (dataGridViewCount.Columns.Count > 0)
                dataGridViewCount.Columns[0].Visible = false;


        }


        private void listViewMain_Click(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection indexes = this.listViewMain.SelectedIndices;
            intMenu = indexes[0];

            if (intMenu == 6)
            {
                this.Dispose();
                return;
            }

            if (intMenu != 5) //ϵͳ����
            {
                if (intMenu == 2) //ǩ��
                {
                    if (strConn == "") //�����ݿ�����
                    {
                        if (MessageBox.Show("���ݿ�����δ���ã��Ƿ��������ǩ����", "���ݿ����", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                        {
                            formoffSign frmoffSign = new formoffSign();
                            frmoffSign.strVersion = strVersion;
                            frmoffSign.intCardLength = intCardLength;

                            frmoffSign.Show(this);
                            intMenu = 5; intContent = 0; initInteface();
                            return;
                        }
                        intMenu = 5; intContent = 0;
                    }
                    

                }
                else
                {
                    if (strConn == "") //�����ݿ�����
                    {
                        MessageBox.Show("���ݿ�����δ���ã�", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        intMenu = 5; intContent = 0;
                    }
                }
            }
            initInteface();
        }

        private void listViewSet_DoubleClick(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection indexes = this.listViewSet.SelectedIndices;
            intContent = indexes[0];

            switch (intContent)
            {
                case 0: //���ݿ�����
                    formDatabaseSet frmDatabaseSet = new formDatabaseSet();
                    frmDatabaseSet.intMode = 0;//����ģʽ

                    if (frmDatabaseSet.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                    {
                        strConn = frmDatabaseSet.strConn;
                        if (strConn == "") //���Ӵ���
                            return;

                        //��ʼ������
                        sqlConn.ConnectionString = strConn;
                        sqlComm.Connection = sqlConn;
                        sqlDA.SelectCommand = sqlComm;

                        System.Data.SqlClient.SqlDataReader sqldr;
                        sqlComm.CommandText = "SELECT ��˾�� FROM ϵͳ������";
                        sqlConn.Open();
                        sqldr = sqlComm.ExecuteReader();
                        if (sqldr.HasRows)
                        {
                            sqldr.Read();
                            this.Text = "������ҵ�������ϵͳ��"+sqldr.GetValue(0).ToString();
                            //this.Text = "�������ϵͳ��" + sqldr.GetValue(0).ToString();
                            sqldr.Close();
                        }
                        sqlConn.Close();

                        //��Ա����
                        InitializePeopleDataGridView();
                        //�������
                        InitializeConferenceDataGridView();
                        //ǩ������
                        InitializedataGridViewSign();
                        //���ѹ���
                        InitializedataGridViewConsume();
                        //ǩ��ͳ��
                        InitializedataGridViewCount();
                    }
                    break;
                case 1://���ܿ�����
                    formCardSet frmCardSet = new formCardSet();
                    if (frmCardSet.ShowDialog(this) == System.Windows.Forms.DialogResult.Yes)
                    {
                        intCard = frmCardSet.intCard;
                        intComNumber = frmCardSet.intComNumber;
                        intCardLength = frmCardSet.intLength;
                        lComBand = frmCardSet.lComBand;

                    }
                    break;
                case 2: //ϵͳ��������
                    if (strConn == "")
                    {
                        MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    formSystemSet frmSystemSet = new formSystemSet();
                    frmSystemSet.strConn = strConn;
                    frmSystemSet.strVersion = strVersion;
                    frmSystemSet.ShowDialog(this);
                    this.Text = "������ҵ�������ϵͳ��" + frmSystemSet.strCompanyName;
                    //this.Text = "�������ϵͳ��" + frmSystemSet.strCompanyName;
                    break;

                case 3: //��������
                    if (strConn == "")
                    {
                        MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    formDataMange frmDataMange = new formDataMange();
                    frmDataMange.strConn = strConn;
                    frmDataMange.ShowDialog(this);
                    break;
                case 4: //�������ݿ�
                    formDatabaseSet frmDatabaseSet1 = new formDatabaseSet();
                    frmDatabaseSet1.intMode = 1;//����ģʽ

                    frmDatabaseSet1.ShowDialog(this);
                    strConn = frmDatabaseSet1.strConn;
                    if (strConn == "") return;

                    //��ʼ������
                    sqlConn.ConnectionString = strConn;
                    sqlComm.Connection = sqlConn;
                    sqlDA.SelectCommand = sqlComm;

                    //��Ա����
                    InitializePeopleDataGridView();
                    //�������
                    InitializeConferenceDataGridView();
                    //ǩ������
                    InitializedataGridViewSign();
                    //���ѹ���
                    InitializedataGridViewConsume();
                    //ǩ��ͳ��
                    InitializedataGridViewCount();
                    break;

                default:
                    break;
            }
        }
        private void InitializedataGridViewConsume()
        {
            DataTable DataTableConference = dSet.Tables["CONFERENCE"];
            dataGridViewConsume.DataSource = DataTableConference;

            dataGridViewConsume.Columns[0].Visible = false;

            dataGridViewConsume.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewConsume.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewConsume.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewConsume.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private void InitializedataGridViewCount()
        {
            DataTable DataTableConference = dSet.Tables["CONFERENCE"];
            dataGridViewCount.DataSource = DataTableConference;

            dataGridViewCount.Columns[0].Visible = false;

            dataGridViewCount.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCount.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCount.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCount.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void InitializedataGridViewSign()
        {
                DataTable DataTableConference = dSet.Tables["CONFERENCE"];
                dataGridViewSign.DataSource = DataTableConference;

                dataGridViewSign.Columns[0].Visible = false;

                dataGridViewSign.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewSign.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewSign.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private void InitializeConferenceDataGridView()
        {
            try
            {

                sqlComm.CommandText = "SELECT ����ID, ��������, ���鸱��, ��ʼʱ�� FROM ����� ORDER BY ��ʼʱ�� DESC";
                DataTable DataTableConference;

                sqlConn.Open();

                if (dSet.Tables.Contains("CONFERENCE")) dSet.Tables.Remove("CONFERENCE");
                sqlDA.Fill(dSet, "CONFERENCE");

                DataTableConference = dSet.Tables["CONFERENCE"];
                dataGridViewConference.DataSource = DataTableConference;


                //dataGridViewConference.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                //dataGridViewConference.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridViewConference.Columns[0].Visible = false;

                dataGridViewConference.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewConference.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewConference.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewConference.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


                sqlConn.Close();
            }
            catch (Exception SqlException)
            {
                MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Threading.Thread.CurrentThread.Abort();
            }


        }

        private void InitializePeopleDataGridView()
        {
            try
            {

                
                sqlComm.CommandText = strPSelect;
                DataTable DataTablePeople ;

                sqlConn.Open();

                if (dSet.Tables.Contains("PEOPLE")) dSet.Tables.Remove("PEOPLE");
                sqlDA.Fill(dSet, "PEOPLE");

                DataTablePeople = dSet.Tables["PEOPLE"];
                dataGridViewPeople.DataSource = DataTablePeople;

                
                //dataGridViewPeople.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                //dataGridViewPeople.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridViewPeople.Columns[0].Visible = false;

                dataGridViewPeople.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewPeople.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                sqlConn.Close();
            }
            catch (Exception SqlException)
            {
                MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Threading.Thread.CurrentThread.Abort();
            }


        }

        private void btnAddPeople_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��","���ݿ����",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            formPeopleAdd frmPeopleAdd = new formPeopleAdd(1);

            frmPeopleAdd.strConn = strConn; //����ģʽ
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.bAddCon = false;


            frmPeopleAdd.ShowDialog(this);

            //ˢ����Ա�б�
            strPSelect = "SELECT ��ԱID, ����, �Ա�, ������λ, ְ��, ְ��, ͨѶ��ַ, ��Ա���, ������ FROM ��Ա��";
            InitializePeopleDataGridView();
        }

        private void btnDelPeople_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("��ѡ��Ҫɾ������Ա��Ϣ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�Ƿ�Ҫɾ����ѡ������Ա��Ϣ����������޷��ָ���", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //ɾ��ѡ������Ա
            sqlConn.Open();

            int i=0;
            int intPID;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            try
            {
                for (i = 0; i < dataGridViewPeople.SelectedRows.Count; i++)
                {
                    intPID = Int32.Parse(dataGridViewPeople.SelectedRows[i].Cells[0].Value.ToString());
                    sqlComm.CommandText = "DELETE FROM ��Ա�� WHERE (��ԱID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM ��Ƭ�� WHERE (��ԱID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                }
                sqlta.Commit();
            }
            catch (Exception err)
            {
                //�ع�
                MessageBox.Show("���ݿ����", "���ݿ����:"+err.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("�ɹ�ɾ��" + i.ToString() + "����Ա��Ϣ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            sqlConn.Close();
            //ˢ����Ա�б�
            InitializePeopleDataGridView();
        }

        private void btnEditPeople_Click(object sender, EventArgs e)
        {
            int intSelectIndex;

            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value==null))
            {
                MessageBox.Show("��ѡ��Ҫ�޸ĵ���Ա��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            formPeopleAdd frmPeopleAdd = new formPeopleAdd(2);
            frmPeopleAdd.strConn = strConn; //�޸�ģʽ
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;

            intSelectIndex = dataGridViewPeople.SelectedRows[0].Index;
            frmPeopleAdd.intPID=Int32.Parse(dataGridViewPeople.SelectedRows[0].Cells[0].Value.ToString());
            frmPeopleAdd.ShowDialog(this);

           
            //ˢ����Ա�б�
            InitializePeopleDataGridView();
            dataGridViewPeople.Rows[0].Selected = false;
            dataGridViewPeople.Rows[intSelectIndex].Selected = true;
            dataGridViewPeople.FirstDisplayedScrollingRowIndex = intSelectIndex;


        }

        private void btnInputPeople_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formExcelInput frmExcelInput = new formExcelInput();
            frmExcelInput.strConn = strConn; //�޸�ģʽ
            frmExcelInput.strVersion = strVersion;
            frmExcelInput.ShowDialog(this);
            //ˢ����Ա�б�
            InitializePeopleDataGridView();
        }

        private void mainCon_Load(object sender, EventArgs e)
        {
            //��ʼ������
            toolStripStatusLabelMain.Alignment = ToolStripItemAlignment.Left;
            this.timerClock.Start();
            toolStripStatusLabelS.Alignment = ToolStripItemAlignment.Right;

            //��ʼ������Ϣ
            string dFileName = "";
            dFileName = Directory.GetCurrentDirectory() + "\\cardcon.xml";
            if (File.Exists(dFileName)) //�����ļ�
            {
                dSet.ReadXml(dFileName);
                intCardLength = Int32.Parse(dSet.Tables["���ܿ���Ϣ"].Rows[0][2].ToString());

                intCardLength = Int32.Parse(dSet.Tables["���ܿ���Ϣ"].Rows[0][2].ToString());
                intComNumber = Int32.Parse(dSet.Tables["���ܿ���Ϣ"].Rows[0][1].ToString());
                lComBand = Int32.Parse(dSet.Tables["���ܿ���Ϣ"].Rows[0][3].ToString());
                intCard = Int32.Parse(dSet.Tables["���ܿ���Ϣ"].Rows[0][0].ToString());

                dSet.Tables.Clear();
            }
            else  //��ʼ��
            {
                intCard = 1;
                intComNumber = 1;
                intCardLength = 10;
                lComBand = 19200;
            }



            //���ע����Ϣ
            //if (checkSoftSerial()) //����ע��
            //{
                if (strConn == "")
                {
                    intMenu = 5; intContent = 0;
                    initInteface();
                    connDataBase();
                    return;
                }
            //}
            //else //û��ע��
            //{
            //    MessageBox.Show("δ�����������������������", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    intMenu = 5; intContent = 0;
            //    initInteface();
            //    connDataBase();
            //    //strVersion = "-1";
            //    this.Close();
            //}
        }

        private void connDataBase()
        {
            string dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
            if (File.Exists(dFileName)) //�����ļ�
            {
                dSet.ReadXml(dFileName);
            }
            else
                return;

            strConn = "workstation id=CY;packet size=4096;user id=" + dSet.Tables["���ݿ���Ϣ"].Rows[0][1].ToString() + ";password=" + dSet.Tables["���ݿ���Ϣ"].Rows[0][2].ToString() + ";data source=\"" + dSet.Tables["���ݿ���Ϣ"].Rows[0][0].ToString() + "\";;initial catalog=" + dSet.Tables["���ݿ���Ϣ"].Rows[0][3].ToString();
            dSet.Tables.Clear();
            

            sqlConn.ConnectionString = strConn;
            try
            {
                sqlConn.Open();
            }
            catch (System.Data.SqlClient.SqlException err)
            {
                MessageBox.Show("���ݿ����Ӵ����������Ա��ϵ");
                strConn = "";
                return;

            }
            sqlConn.Close();

            //��ʼ������
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            //��Ա����
            InitializePeopleDataGridView();
            //�������
            InitializeConferenceDataGridView();
            //ǩ������
            InitializedataGridViewSign();
            //���ѹ���
            InitializedataGridViewConsume();
            //ǩ��ͳ��
            InitializedataGridViewCount();

        }

        private void dataGridViewPeople_DoubleClick(object sender, EventArgs e)
        {
            btnEditPeople_Click(null, null);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();

            frmConference.strConn = strConn;

            frmConference.intMode = 1; //����ģʽ
            frmConference.strVersion = strVersion;

            frmConference.ShowDialog(this);

            //ˢ�»����б�
            InitializeConferenceDataGridView();
            //ǩ������
            InitializedataGridViewSign();
            //���ѹ���
            InitializedataGridViewConsume();
            //ǩ��ͳ��
            InitializedataGridViewCount();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewConference.SelectedRows.Count == 0) || (dataGridViewConference.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("��ѡ��Ҫɾ���Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�Ƿ�Ҫɾ����ѡ���Ļ�����Ϣ����������޷��ָ���", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //ɾ��ѡ���Ļ���
            sqlConn.Open();

            int i=0;
            int intCID;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            try
            {
                for (i = 0; i < dataGridViewConference.SelectedRows.Count; i++)
                {
                    intCID = Int32.Parse(dataGridViewConference.SelectedRows[i].Cells[0].Value.ToString());
                    sqlComm.CommandText = "DELETE FROM ����� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM ����� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM �������ѱ� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM �λ���Ա�� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM ǩ���� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
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
                MessageBox.Show("�ɹ�ɾ��" + i.ToString() + "��������Ϣ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            sqlConn.Close();

            //ˢ�»����б�
            InitializeConferenceDataGridView();
            //ǩ������
            InitializedataGridViewSign();
            //���ѹ���
            InitializedataGridViewConsume();
            //ǩ��ͳ��
            InitializedataGridViewCount();
            
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewConference.SelectedRows.Count<1)
            {
                MessageBox.Show("û��ѡ��Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();

            frmConference.strConn = strConn;
            frmConference.strVersion = strVersion;
            frmConference.intMode = 2; //�޸�ģʽ
            frmConference.intConferenceID = Int32.Parse(dataGridViewConference.SelectedRows[0].Cells[0].Value.ToString());

            frmConference.ShowDialog(this);

            //ˢ�»����б�
            InitializeConferenceDataGridView();
            //ǩ������
            InitializedataGridViewSign();
            //���ѹ���
            InitializedataGridViewConsume();
            //ǩ��ͳ��
            InitializedataGridViewCount();
        }

        private void dataGridViewConference_DoubleClick(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();

            frmConference.strConn = strConn;
            frmConference.strVersion = strVersion;
            frmConference.intMode = 2; //�޸�ģʽ
            frmConference.intConferenceID = Int32.Parse(dataGridViewConference.SelectedRows[0].Cells[0].Value.ToString());

            frmConference.ShowDialog(this);

            //ˢ�»����б�
            InitializeConferenceDataGridView();
        }

        private void btnSign_Click(object sender, EventArgs e)
        {

            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewSign.SelectedRows.Count < 1)
            {
                MessageBox.Show("û��ѡ��Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formSign frmSign = new formSign();

            frmSign.strConn = strConn;
            frmSign.strVersion = strVersion;
            frmSign.intCardLength = intCardLength;
            frmSign.intCard = intCard;
            frmSign.intComNumber = intComNumber;
            frmSign.lComBand = lComBand;

            frmSign.intConferenceID = Int32.Parse(dataGridViewSign.SelectedRows[0].Cells[0].Value.ToString());

            frmSign.ShowDialog(this);

            //ˢ�»����б�
            //InitializeConferenceDataGridView();

        }

        private void btnResign_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewSign.SelectedRows.Count < 1)
            {
                MessageBox.Show("û��ѡ��Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formreSign frmreSign = new formreSign();

            frmreSign.strConn = strConn;
            frmreSign.intConferenceID = Int32.Parse(dataGridViewSign.SelectedRows[0].Cells[0].Value.ToString());

            frmreSign.ShowDialog(this);
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
            ͳ�����ToolStripMenuItem_Click(null, null);
        }

        private void btnCountAll_Click(object sender, EventArgs e)
        {
            ͳ��ȫ��ToolStripMenuItem_Click(null, null);
        }

        private void btnConsumeDef_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewConsume.SelectedRows.Count < 1)
            {
                MessageBox.Show("û��ѡ��Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConsumeDef frmConsumeDef = new formConsumeDef();

            frmConsumeDef.strConn = strConn;
            frmConsumeDef.strVersion = strVersion;

            frmConsumeDef.intConferenceID = Int32.Parse(dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString());

            frmConsumeDef.ShowDialog(this);
        }

        private void btnConsume_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewConsume.SelectedRows.Count < 1)
            {
                MessageBox.Show("û��ѡ��Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConsumeSign frmConsumeSign = new formConsumeSign();

            frmConsumeSign.strConn = strConn;
            frmConsumeSign.strVersion = strVersion;
            frmConsumeSign.intCardLength = intCardLength;
            frmConsumeSign.intCard = intCard;
            frmConsumeSign.intComNumber = intComNumber;
            frmConsumeSign.lComBand = lComBand;

            frmConsumeSign.intConferenceID = Int32.Parse(dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString());

            frmConsumeSign.ShowDialog(this);
        }

        private void ����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBoxIT frmAbout = new AboutBoxIT();

            frmAbout.strVersion = strVersion;
            frmAbout.ShowDialog(this);
        }

        private void timerClock_Tick(object sender, EventArgs e)
        {
            this.toolStripStatusLabelClock.Text = System.DateTime.Now.ToLongDateString()+ " " + System.DateTime.Now.ToLongTimeString();
        }

        private void mainCon_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.timerClock.Stop();
        }

        private void ��Ա����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            intMenu = 0;
            initInteface();
        }

        private void �������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 1;
            initInteface();
        }

        private void ���ѹ���ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 3;
            initInteface();
        }

        private void ������ԱToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 0;
            initInteface();

            formPeopleAdd frmPeopleAdd = new formPeopleAdd(1);
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.strConn = strConn; //����ģʽ
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.ShowDialog(this);

            //ˢ����Ա�б�
            InitializePeopleDataGridView();
        }

        private void DelPeopleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("��ѡ��Ҫɾ������Ա��Ϣ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�Ƿ�Ҫɾ����ѡ������Ա��Ϣ����������޷��ָ���", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //ɾ��ѡ������Ա
            sqlConn.Open();

            int i=0;
            int intPID;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            try
            {
                for (i = 0; i < dataGridViewPeople.SelectedRows.Count; i++)
                {
                    intPID = Int32.Parse(dataGridViewPeople.SelectedRows[i].Cells[0].Value.ToString());
                    sqlComm.CommandText = "DELETE FROM ��Ա�� WHERE (��ԱID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM ��Ƭ�� WHERE (��ԱID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                }
                sqlta.Commit();
            }
            catch (Exception err)
            {
                //�ع�
                MessageBox.Show("���ݿ����", "���ݿ����:" + err.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("�ɹ�ɾ��" + i.ToString() + "����Ա��Ϣ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            sqlConn.Close();
            //ˢ����Ա�б�
            InitializePeopleDataGridView();

        }

        private void ��Ա�༭ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            intMenu = 0;
            initInteface();

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("��ѡ��Ҫ�޸ĵ���Ա��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            formPeopleAdd frmPeopleAdd = new formPeopleAdd(2);

            frmPeopleAdd.strConn = strConn; //�޸�ģʽ
            frmPeopleAdd.intPID = Int32.Parse(dataGridViewPeople.SelectedRows[0].Cells[0].Value.ToString());
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.ShowDialog(this);

            //ˢ����Ա�б�
            InitializePeopleDataGridView();
        }

        private void excel����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 0;
            initInteface();

            formExcelInput frmExcelInput = new formExcelInput();
            frmExcelInput.strConn = strConn;
            frmExcelInput.strVersion = strVersion;
            if (frmExcelInput.ShowDialog(this) != DialogResult.Cancel)
                //ˢ����Ա�б�
                InitializePeopleDataGridView();
        }

        private void ���ӻ���ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 1;
            initInteface();

            formConference frmConference = new formConference();
            frmConference.strConn = strConn;
            frmConference.strVersion = strVersion;
            frmConference.intMode = 1; //����ģʽ

            frmConference.ShowDialog(this);
            //ˢ�»����б�
            InitializeConferenceDataGridView();
        }

        private void �޸Ļ���ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();
            intMenu = 1;
            initInteface();

            frmConference.strConn = strConn;
            frmConference.intMode = 2; //�޸�ģʽ
            frmConference.strVersion = strVersion;
            frmConference.intConferenceID = Int32.Parse(dataGridViewConference.SelectedRows[0].Cells[0].Value.ToString());

            frmConference.ShowDialog(this);

            //ˢ�»����б�
            InitializeConferenceDataGridView();
        }

        private void DelConferenceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewConference.SelectedRows.Count == 0) || (dataGridViewConference.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("��ѡ��Ҫɾ���Ļ��飡", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�Ƿ�Ҫɾ����ѡ���Ļ�����Ϣ����������޷��ָ���", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //ɾ��ѡ���Ļ���
            sqlConn.Open();

            int i=0;
            int intCID;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            try
            {
                for (i = 0; i < dataGridViewPeople.SelectedRows.Count; i++)
                {
                    intCID = Int32.Parse(dataGridViewPeople.SelectedRows[i].Cells[0].Value.ToString());
                    sqlComm.CommandText = "DELETE FROM ����� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM ����� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM �������ѱ� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM �λ���Ա�� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM ǩ���� WHERE (����ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
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
                MessageBox.Show("�ɹ�ɾ��" + i.ToString() + "��������Ϣ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            sqlConn.Close();

            //ˢ�»����б�
            InitializeConferenceDataGridView();

        }

        private void ���Ѷ���ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 3;
            initInteface();

            formConsumeDef frmConsumeDef = new formConsumeDef();

            frmConsumeDef.strConn = strConn;
            frmConsumeDef.strVersion = strVersion;
            frmConsumeDef.intConferenceID = Int32.Parse(dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString());

            frmConsumeDef.ShowDialog(this);
        }

        private void ��������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 3;
            initInteface();
            formConsumeSign frmConsumeSign = new formConsumeSign();

            frmConsumeSign.strConn = strConn;
            frmConsumeSign.strVersion = strVersion;
            frmConsumeSign.intConferenceID = Int32.Parse(dataGridViewConsume.SelectedRows[0].Cells[0].Value.ToString());

            frmConsumeSign.ShowDialog(this);

        }

        private void �˳�ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void ����ǩ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 2;
            initInteface();
            formSign frmSign = new formSign();

            frmSign.strConn = strConn;
            frmSign.strVersion = strVersion;
            frmSign.intCardLength = intCardLength;
            frmSign.intCard = intCard;
            frmSign.intComNumber = intComNumber;
            frmSign.lComBand = lComBand;
            frmSign.intConferenceID = Int32.Parse(dataGridViewSign.SelectedRows[0].Cells[0].Value.ToString());

            frmSign.ShowDialog(this);

            //ˢ�»����б�
            //InitializeConferenceDataGridView();
        }

        private void ������ǩToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 2;
            initInteface();
            formreSign frmreSign = new formreSign();

            frmreSign.strConn = strConn;
            frmreSign.intConferenceID = Int32.Parse(dataGridViewSign.SelectedRows[0].Cells[0].Value.ToString());

            frmreSign.ShowDialog(this);
        }

        private void ����ǩ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 2;
            initInteface();
            formoffSign frmoffSign = new formoffSign();
            frmoffSign.strVersion = strVersion;
            frmoffSign.intCardLength = intCardLength;
            frmoffSign.intCard = intCard;
            frmoffSign.intComNumber = intComNumber;
            frmoffSign.lComBand = lComBand;
            frmoffSign.Show(this);
        }

        private void ͳ��ȫ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 4;
            initInteface();
            formCount frmCount = new formCount();

            frmCount.strConn = strConn;
            frmCount.strVersion = strVersion;
            frmCount.intConferenceID = 0;
            frmCount.intCardLength = intCardLength;
            frmCount.intCard = intCard;
            frmCount.intComNumber = intComNumber;
            frmCount.lComBand = lComBand;

            frmCount.ShowDialog(this);
        }

        private void ͳ�����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 4;
            initInteface();
            formCount frmCount = new formCount();

            frmCount.strConn = strConn;
            frmCount.strVersion = strVersion;
            frmCount.intConferenceID = Int32.Parse(dataGridViewCount.SelectedRows[0].Cells[0].Value.ToString());
            frmCount.intCardLength = intCardLength;
            frmCount.intCard = intCard;
            frmCount.intComNumber = intComNumber;
            frmCount.lComBand = lComBand;

            frmCount.ShowDialog(this);
        }

        private void ���ݿ�����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();

            formDatabaseSet frmDatabaseSet = new formDatabaseSet();
            frmDatabaseSet.intMode = 0;//����ģʽ

            if (frmDatabaseSet.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                strConn = frmDatabaseSet.strConn;
                if (strConn == "") //���Ӵ���
                    return;

                //��ʼ������
                sqlConn.ConnectionString = strConn;
                sqlComm.Connection = sqlConn;
                sqlDA.SelectCommand = sqlComm;

                //��Ա����
                InitializePeopleDataGridView();
                //�������
                InitializeConferenceDataGridView();
                //ǩ������
                InitializedataGridViewSign();
                //���ѹ���
                InitializedataGridViewConsume();
                //ǩ��ͳ��
                InitializedataGridViewCount();
            }
        }

        private void ���ܿ�����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();
            
            formCardSet frmCardSet = new formCardSet();
            if (frmCardSet.ShowDialog(this) == System.Windows.Forms.DialogResult.Yes)
            {
                intCard = frmCardSet.intCard;
                intComNumber = frmCardSet.intComNumber;
                intCardLength = frmCardSet.intLength;
                lComBand = frmCardSet.lComBand;
            }
        }

        private void ϵͳ��������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();

            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formSystemSet frmSystemSet = new formSystemSet();
            frmSystemSet.strVersion = strVersion;
            frmSystemSet.strConn = strConn;
            frmSystemSet.ShowDialog(this);
        }

        private void ���ݹ���ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();
            
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formDataMange frmDataMange = new formDataMange();
            frmDataMange.strConn = strConn;
            frmDataMange.ShowDialog(this);
        }

        private void �������ݿ�DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();

            formDatabaseSet frmDatabaseSet1 = new formDatabaseSet();
            frmDatabaseSet1.intMode = 1;//����ģʽ

            frmDatabaseSet1.ShowDialog(this);
            strConn = frmDatabaseSet1.strConn;
            if (strConn == "") return;

            //��ʼ������
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            //��Ա����
            InitializePeopleDataGridView();
            //�������
            InitializeConferenceDataGridView();
            //ǩ������
            InitializedataGridViewSign();
            //���ѹ���
            InitializedataGridViewConsume();
            //ǩ��ͳ��
            InitializedataGridViewCount();

        }

        private void btnSearchPeople_Click(object sender, EventArgs e)
        {
            formPeopleSearch frmPeopleSelect = new formPeopleSearch();
            frmPeopleSelect.strConn = strConn; //����ģʽ
            frmPeopleSelect.intCardLength = intCardLength;
            frmPeopleSelect.intCard = intCard;
            frmPeopleSelect.intComNumber = intComNumber;
            frmPeopleSelect.lComBand = lComBand;
            
            frmPeopleSelect.ShowDialog();

            if (frmPeopleSelect.intReturn == 0) return;

            //ѡ��ȫ��
            if (frmPeopleSelect.intReturn == -1) 
                strPSelect = "SELECT ��ԱID, ����, �Ա�, ������λ, ְ��, ְ��, ͨѶ��ַ, ��Ա���, ������ FROM ��Ա��";


            if (frmPeopleSelect.intReturn == 1)
            {
                bool iPosition = true;
                strPSelect = "SELECT ��ԱID, ����, �Ա�, ������λ, ְ��, ְ��, ͨѶ��ַ, ��Ա���, ������ FROM ��Ա�� ";
                if (frmPeopleSelect.textBoxUsername.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (���� LIKE N'%"+frmPeopleSelect.textBoxUsername.Text.Trim()+"%')";
                        iPosition=false;
                    }
                    else
                        strPSelect = strPSelect + " AND (���� LIKE N'%" + frmPeopleSelect.textBoxUsername.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.comboBoxGender.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (�Ա� LIKE N'%" + frmPeopleSelect.comboBoxGender.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (�Ա� LIKE N'%" + frmPeopleSelect.comboBoxGender.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxCompany.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (������λ LIKE N'%" + frmPeopleSelect.textBoxCompany.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (������λ LIKE N'%" + frmPeopleSelect.textBoxCompany.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxPPost.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (ְ�� LIKE N'%" + frmPeopleSelect.textBoxPPost.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (ְ�� LIKE N'%" + frmPeopleSelect.textBoxPPost.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxPost.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (ְ�� LIKE N'%" + frmPeopleSelect.textBoxPost.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (ְ�� LIKE N'%" + frmPeopleSelect.textBoxPost.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxAddress.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (ͨѶ��ַ LIKE N'%" + frmPeopleSelect.textBoxAddress.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (ͨѶ��ַ LIKE N'%" + frmPeopleSelect.textBoxAddress.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.comboBoxGroup.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (��Ա��� = N'" + frmPeopleSelect.comboBoxGroup.Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (��Ա��� = N'%" + frmPeopleSelect.comboBoxGroup.Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxCard.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (������ = N'" + frmPeopleSelect.textBoxCard.Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (������ = N'%" + frmPeopleSelect.textBoxCard.Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[0].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (��������һ = N'" + frmPeopleSelect.textBoxAddi[0].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (��������һ = N'%" + frmPeopleSelect.textBoxAddi[0].Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[1].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (�������Զ� = N'" + frmPeopleSelect.textBoxAddi[1].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (�������Զ� = N'%" + frmPeopleSelect.textBoxAddi[1].Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[2].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (���������� = N'" + frmPeopleSelect.textBoxAddi[2].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (���������� = N'%" + frmPeopleSelect.textBoxAddi[2].Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[3].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (���������� = N'" + frmPeopleSelect.textBoxAddi[3].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (���������� = N'%" + frmPeopleSelect.textBoxAddi[3].Text.Trim() + "')";
                }



            }

            InitializePeopleDataGridView();
        }

        private void btnShowAll_Click(object sender, EventArgs e)
        {
            strPSelect = "SELECT ��ԱID, ����, �Ա�, ������λ, ְ��, ְ��, ͨѶ��ַ, ��Ա���, ������ FROM ��Ա��";
            InitializePeopleDataGridView();
        }

        private void ��ԱToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formPeopleAdd frmPeopleAdd = new formPeopleAdd(1);

            frmPeopleAdd.strConn = strConn; //����ģʽ
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.bAddCon = true;


            frmPeopleAdd.ShowDialog(this);

            //ˢ����Ա�б�
            strPSelect = "SELECT ��ԱID, ����, �Ա�, ������λ, ְ��, ְ��, ͨѶ��ַ, ��Ա���, ������ FROM ��Ա��";
            InitializePeopleDataGridView();

        }

        private void ��Ա��ѯToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSearchPeople_Click(null, null);
        }

        private bool checkSoftSerial()
        {
            bool boolSign = true;

            ClassSenseLock cSenseLock = new ClassSenseLock();

            

            cSenseLock.strModal = "itconference";
            cSenseLock.iStart = 0;
            cSenseLock.iLength = 16;

            int iR = cSenseLock.ReadSenseLock();

            if (iR != 0)
                boolSign = false;
            else
                strVersion = cSenseLock.sVersion;



            /*
            string dFileName = Directory.GetCurrentDirectory() + "\\serial.xml";
            if (File.Exists(dFileName)) //�����ļ�
            {
                dSet.ReadXml(dFileName);
                strSerial = dSet.Tables["���ע��"].Rows[0][0].ToString();
                dSet.Tables.Clear();

                chrislock lockit=new chrislock();

                string strData = ""; //ԭʼ��
                strData = lockit.getMacAddress();
                if (strData == "")
                    strData = lockit.getDiskSerial();
                if (strData == "")
                {
                    strData = "CHRISTIE(CHENYI)";
                }

                string strEnData = strSerial;
                string strtemp;
                string strpassword = "christie";
                strtemp = lockit.getUnlockEnData(strEnData, strpassword);
                string[] strArray = strtemp.Split(new char[1] { '#' });
                if (strArray.Length < 3)
                {
                    MessageBox.Show("�����δ����ע�᣿", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    boolSign = false;
                }
                else
                {
                    strVersion = strArray[2];
                    if (strArray[1] == strData || strArray[1] == "CHRISTIE(CHENYI)")
                    {
                        if (strArray[0] != "CHRISTIE(CHENYI)") //��������
                        {
                            try
                            {
                                System.DateTime dtRegister=System.DateTime.Parse(strArray[0]);
                                if(dtRegister < System.DateTime.Now)
                                {
                                    MessageBox.Show("���ע����Ϣ�ѹ��ڣ�", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    boolSign = false;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("���ע����Ϣ����ȷ��", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                boolSign = false;
                            }
                            
                        }
                    }
                    else //��֤�벻��ȷ
                    {
                        MessageBox.Show("���ע����Ϣ����ȷ��", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        boolSign = false;
                    }
                }

            }
            else  //û��ע��
            {
                strSerial = "";
                MessageBox.Show("�����δ����ע�᣿", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                boolSign = false;
            }

            if (boolSign)
            {
                //
                listViewMain.Enabled = true;
                ����MToolStripMenuItem.Enabled = true;
                ǩ��ToolStripMenuItem.Enabled = true;
                ͳ��ToolStripMenuItem.Enabled = true;
                ����ToolStripMenuItem.Enabled = true;
            }
            else
            {
                listViewMain.Enabled = false;
                ����MToolStripMenuItem.Enabled = false;
                ǩ��ToolStripMenuItem.Enabled = false;
                ͳ��ToolStripMenuItem.Enabled = false;
                ����ToolStripMenuItem.Enabled = false;
            }
             */ 

            return boolSign;
        }

        private void ���ע��RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formSoftSign frmSoftSign = new formSoftSign();
            frmSoftSign.ShowDialog();
          
            if (frmSoftSign.dlResult != DialogResult.Cancel)
            {
                MessageBox.Show("���ע�����,���������л������ϵͳ��", "ע��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Dispose();
            }
            
        }

        private void �������Թ���TToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 0;
            initInteface();

            formPAddiMange frmPAddiMange = new formPAddiMange();
            frmPAddiMange.strConn = strConn;
            if(frmPAddiMange.ShowDialog()!=DialogResult.Cancel)
                InitializePeopleDataGridView();

        }

        private void ��λ����DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            FormDW frmDW = new FormDW();

            frmDW.strConn = strConn;
            frmDW.strVersion = strVersion;

            frmDW.ShowDialog(this);
        }

        private void ��ѵToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formTrainManage frmTrainManage = new formTrainManage();

            frmTrainManage.strConn = strConn;
            frmTrainManage.intMod = 0;
            frmTrainManage.strVersion = strVersion;

            frmTrainManage.ShowDialog(this);
        }

        private void ��ѵǩ��SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            formTrainManage frmTrainManage = new formTrainManage();

            frmTrainManage.strConn = strConn;
            frmTrainManage.intMod = 1;
            frmTrainManage.strVersion = strVersion;

            frmTrainManage.intCardLength = intCardLength;
            frmTrainManage.intCard = intCard;
            frmTrainManage.intComNumber = intComNumber;
            frmTrainManage.lComBand = lComBand;

            frmTrainManage.ShowDialog(this);
        }

        private void ��ѵͳ��CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formTrainManage frmTrainManage = new formTrainManage();

            frmTrainManage.strConn = strConn;
            frmTrainManage.intMod = 2;
            frmTrainManage.strVersion = strVersion;

            frmTrainManage.ShowDialog(this);
        }

        private void ����Ŀ¼CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, "itConference.chm");   
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            string strT = "��Ա����";
            PrintDGV.Print_DataGridView(dataGridViewPeople, strT, false);
        }

        private void �ֳֻ�ǩ����ϢToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("���ݿ����ò���ȷ��", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formSCJ frmSCJ = new formSCJ();
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

            formSCJQD frmSCJQD = new formSCJQD();

            frmSCJQD.strConn = strConn;
            frmSCJQD.strVersion = strVersion;
            frmSCJQD.intCardLength = intCardLength;
            frmSCJQD.intCard = intCard;
            frmSCJQD.intComNumber = intComNumber;
            frmSCJQD.lComBand = lComBand;

            frmSCJQD.ShowDialog(this);

        }



    }
}