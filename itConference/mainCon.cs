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
            
            //初始化

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
            strPSelect = "SELECT 人员ID, 姓名, 性别, 所属单位, 职称, 职务, 通讯地址, 人员类别, 主卡号 FROM 人员表";

            

        }

        private void initInteface()
        {
            int i;

            if (intMenu <  listViewMain.Items.Count-1) //非退出
                toolStripStatusLabelS.Text = listViewMain.Items[intMenu].Text;

            if (intMenu != 0) //删除人员菜单
                DelPeopleToolStripMenuItem.Enabled = false;
            else
                DelPeopleToolStripMenuItem.Enabled = true;
            

            if (intMenu != 1) //删除会议菜单
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

            if (intMenu != 5) //系统设置
            {
                if (intMenu == 2) //签到
                {
                    if (strConn == "") //无数据库连接
                    {
                        if (MessageBox.Show("数据库连接未设置！是否进行离线签到？", "数据库错误", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
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
                    if (strConn == "") //无数据库连接
                    {
                        MessageBox.Show("数据库连接未设置！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                case 0: //数据库设置
                    formDatabaseSet frmDatabaseSet = new formDatabaseSet();
                    frmDatabaseSet.intMode = 0;//测试模式

                    if (frmDatabaseSet.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                    {
                        strConn = frmDatabaseSet.strConn;
                        if (strConn == "") //连接错误
                            return;

                        //初始化窗口
                        sqlConn.ConnectionString = strConn;
                        sqlComm.Connection = sqlConn;
                        sqlDA.SelectCommand = sqlComm;

                        System.Data.SqlClient.SqlDataReader sqldr;
                        sqlComm.CommandText = "SELECT 公司名 FROM 系统参数表";
                        sqlConn.Open();
                        sqldr = sqlComm.ExecuteReader();
                        if (sqldr.HasRows)
                        {
                            sqldr.Read();
                            this.Text = "倚天立业会议管理系统："+sqldr.GetValue(0).ToString();
                            //this.Text = "会议管理系统：" + sqldr.GetValue(0).ToString();
                            sqldr.Close();
                        }
                        sqlConn.Close();

                        //人员管理
                        InitializePeopleDataGridView();
                        //会议管理
                        InitializeConferenceDataGridView();
                        //签到管理
                        InitializedataGridViewSign();
                        //消费管理
                        InitializedataGridViewConsume();
                        //签到统计
                        InitializedataGridViewCount();
                    }
                    break;
                case 1://智能卡设置
                    formCardSet frmCardSet = new formCardSet();
                    if (frmCardSet.ShowDialog(this) == System.Windows.Forms.DialogResult.Yes)
                    {
                        intCard = frmCardSet.intCard;
                        intComNumber = frmCardSet.intComNumber;
                        intCardLength = frmCardSet.intLength;
                        lComBand = frmCardSet.lComBand;

                    }
                    break;
                case 2: //系统参数设置
                    if (strConn == "")
                    {
                        MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    formSystemSet frmSystemSet = new formSystemSet();
                    frmSystemSet.strConn = strConn;
                    frmSystemSet.strVersion = strVersion;
                    frmSystemSet.ShowDialog(this);
                    this.Text = "倚天立业会议管理系统：" + frmSystemSet.strCompanyName;
                    //this.Text = "会议管理系统：" + frmSystemSet.strCompanyName;
                    break;

                case 3: //数据清理
                    if (strConn == "")
                    {
                        MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    formDataMange frmDataMange = new formDataMange();
                    frmDataMange.strConn = strConn;
                    frmDataMange.ShowDialog(this);
                    break;
                case 4: //创建数据库
                    formDatabaseSet frmDatabaseSet1 = new formDatabaseSet();
                    frmDatabaseSet1.intMode = 1;//创建模式

                    frmDatabaseSet1.ShowDialog(this);
                    strConn = frmDatabaseSet1.strConn;
                    if (strConn == "") return;

                    //初始化窗口
                    sqlConn.ConnectionString = strConn;
                    sqlComm.Connection = sqlConn;
                    sqlDA.SelectCommand = sqlComm;

                    //人员管理
                    InitializePeopleDataGridView();
                    //会议管理
                    InitializeConferenceDataGridView();
                    //签到管理
                    InitializedataGridViewSign();
                    //消费管理
                    InitializedataGridViewConsume();
                    //签到统计
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

                sqlComm.CommandText = "SELECT 会议ID, 会议主题, 会议副题, 开始时间 FROM 会议表 ORDER BY 开始时间 DESC";
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
                MessageBox.Show("数据库错误", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                MessageBox.Show("数据库错误", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Threading.Thread.CurrentThread.Abort();
            }


        }

        private void btnAddPeople_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！","数据库错误",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            formPeopleAdd frmPeopleAdd = new formPeopleAdd(1);

            frmPeopleAdd.strConn = strConn; //增加模式
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.bAddCon = false;


            frmPeopleAdd.ShowDialog(this);

            //刷新人员列表
            strPSelect = "SELECT 人员ID, 姓名, 性别, 所属单位, 职称, 职务, 通讯地址, 人员类别, 主卡号 FROM 人员表";
            InitializePeopleDataGridView();
        }

        private void btnDelPeople_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("请选择要删除的人员信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("是否要删除的选定的人员信息？这个操作无法恢复！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //删除选定的人员
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
                    sqlComm.CommandText = "DELETE FROM 人员表 WHERE (人员ID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 照片表 WHERE (人员ID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                }
                sqlta.Commit();
            }
            catch (Exception err)
            {
                //回滚
                MessageBox.Show("数据库错误！", "数据库错误:"+err.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("成功删除" + i.ToString() + "条人员信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            sqlConn.Close();
            //刷新人员列表
            InitializePeopleDataGridView();
        }

        private void btnEditPeople_Click(object sender, EventArgs e)
        {
            int intSelectIndex;

            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value==null))
            {
                MessageBox.Show("请选择要修改的人员！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            formPeopleAdd frmPeopleAdd = new formPeopleAdd(2);
            frmPeopleAdd.strConn = strConn; //修改模式
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;

            intSelectIndex = dataGridViewPeople.SelectedRows[0].Index;
            frmPeopleAdd.intPID=Int32.Parse(dataGridViewPeople.SelectedRows[0].Cells[0].Value.ToString());
            frmPeopleAdd.ShowDialog(this);

           
            //刷新人员列表
            InitializePeopleDataGridView();
            dataGridViewPeople.Rows[0].Selected = false;
            dataGridViewPeople.Rows[intSelectIndex].Selected = true;
            dataGridViewPeople.FirstDisplayedScrollingRowIndex = intSelectIndex;


        }

        private void btnInputPeople_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formExcelInput frmExcelInput = new formExcelInput();
            frmExcelInput.strConn = strConn; //修改模式
            frmExcelInput.strVersion = strVersion;
            frmExcelInput.ShowDialog(this);
            //刷新人员列表
            InitializePeopleDataGridView();
        }

        private void mainCon_Load(object sender, EventArgs e)
        {
            //初始化窗口
            toolStripStatusLabelMain.Alignment = ToolStripItemAlignment.Left;
            this.timerClock.Start();
            toolStripStatusLabelS.Alignment = ToolStripItemAlignment.Right;

            //初始化卡信息
            string dFileName = "";
            dFileName = Directory.GetCurrentDirectory() + "\\cardcon.xml";
            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
                intCardLength = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][2].ToString());

                intCardLength = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][2].ToString());
                intComNumber = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][1].ToString());
                lComBand = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][3].ToString());
                intCard = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][0].ToString());

                dSet.Tables.Clear();
            }
            else  //初始化
            {
                intCard = 1;
                intComNumber = 1;
                intCardLength = 10;
                lComBand = 19200;
            }



            //软件注册信息
            //if (checkSoftSerial()) //已有注册
            //{
                if (strConn == "")
                {
                    intMenu = 5; intContent = 0;
                    initInteface();
                    connDataBase();
                    return;
                }
            //}
            //else //没有注册
            //{
            //    MessageBox.Show("未发现软件狗，请插入软件狗。", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
            }
            else
                return;

            strConn = "workstation id=CY;packet size=4096;user id=" + dSet.Tables["数据库信息"].Rows[0][1].ToString() + ";password=" + dSet.Tables["数据库信息"].Rows[0][2].ToString() + ";data source=\"" + dSet.Tables["数据库信息"].Rows[0][0].ToString() + "\";;initial catalog=" + dSet.Tables["数据库信息"].Rows[0][3].ToString();
            dSet.Tables.Clear();
            

            sqlConn.ConnectionString = strConn;
            try
            {
                sqlConn.Open();
            }
            catch (System.Data.SqlClient.SqlException err)
            {
                MessageBox.Show("数据库连接错误，请与管理员联系");
                strConn = "";
                return;

            }
            sqlConn.Close();

            //初始化窗口
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            //人员管理
            InitializePeopleDataGridView();
            //会议管理
            InitializeConferenceDataGridView();
            //签到管理
            InitializedataGridViewSign();
            //消费管理
            InitializedataGridViewConsume();
            //签到统计
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
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();

            frmConference.strConn = strConn;

            frmConference.intMode = 1; //增加模式
            frmConference.strVersion = strVersion;

            frmConference.ShowDialog(this);

            //刷新会议列表
            InitializeConferenceDataGridView();
            //签到管理
            InitializedataGridViewSign();
            //消费管理
            InitializedataGridViewConsume();
            //签到统计
            InitializedataGridViewCount();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewConference.SelectedRows.Count == 0) || (dataGridViewConference.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("请选择要删除的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("是否要删除的选定的会议信息？这个操作无法恢复！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //删除选定的会议
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
                    sqlComm.CommandText = "DELETE FROM 会议表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 分组表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 会议消费表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 参会人员表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 签到表 WHERE (会议ID = " + intCID.ToString() + ")";
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
            //签到管理
            InitializedataGridViewSign();
            //消费管理
            InitializedataGridViewConsume();
            //签到统计
            InitializedataGridViewCount();
            
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewConference.SelectedRows.Count<1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();

            frmConference.strConn = strConn;
            frmConference.strVersion = strVersion;
            frmConference.intMode = 2; //修改模式
            frmConference.intConferenceID = Int32.Parse(dataGridViewConference.SelectedRows[0].Cells[0].Value.ToString());

            frmConference.ShowDialog(this);

            //刷新会议列表
            InitializeConferenceDataGridView();
            //签到管理
            InitializedataGridViewSign();
            //消费管理
            InitializedataGridViewConsume();
            //签到统计
            InitializedataGridViewCount();
        }

        private void dataGridViewConference_DoubleClick(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();

            frmConference.strConn = strConn;
            frmConference.strVersion = strVersion;
            frmConference.intMode = 2; //修改模式
            frmConference.intConferenceID = Int32.Parse(dataGridViewConference.SelectedRows[0].Cells[0].Value.ToString());

            frmConference.ShowDialog(this);

            //刷新会议列表
            InitializeConferenceDataGridView();
        }

        private void btnSign_Click(object sender, EventArgs e)
        {

            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewSign.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            //刷新会议列表
            //InitializeConferenceDataGridView();

        }

        private void btnResign_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewSign.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            统计输出ToolStripMenuItem_Click(null, null);
        }

        private void btnCountAll_Click(object sender, EventArgs e)
        {
            统计全部ToolStripMenuItem_Click(null, null);
        }

        private void btnConsumeDef_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewConsume.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dataGridViewConsume.SelectedRows.Count < 1)
            {
                MessageBox.Show("没有选择的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void 人员管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            intMenu = 0;
            initInteface();
        }

        private void 会议管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 1;
            initInteface();
        }

        private void 消费管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 3;
            initInteface();
        }

        private void 增加人员ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 0;
            initInteface();

            formPeopleAdd frmPeopleAdd = new formPeopleAdd(1);
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.strConn = strConn; //增加模式
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.ShowDialog(this);

            //刷新人员列表
            InitializePeopleDataGridView();
        }

        private void DelPeopleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("请选择要删除的人员信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("是否要删除的选定的人员信息？这个操作无法恢复！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //删除选定的人员
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
                    sqlComm.CommandText = "DELETE FROM 人员表 WHERE (人员ID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 照片表 WHERE (人员ID = " + intPID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                }
                sqlta.Commit();
            }
            catch (Exception err)
            {
                //回滚
                MessageBox.Show("数据库错误！", "数据库错误:" + err.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("成功删除" + i.ToString() + "条人员信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            sqlConn.Close();
            //刷新人员列表
            InitializePeopleDataGridView();

        }

        private void 人员编辑ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            intMenu = 0;
            initInteface();

            if ((dataGridViewPeople.SelectedRows.Count == 0) || (dataGridViewPeople.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("请选择要修改的人员！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            formPeopleAdd frmPeopleAdd = new formPeopleAdd(2);

            frmPeopleAdd.strConn = strConn; //修改模式
            frmPeopleAdd.intPID = Int32.Parse(dataGridViewPeople.SelectedRows[0].Cells[0].Value.ToString());
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.ShowDialog(this);

            //刷新人员列表
            InitializePeopleDataGridView();
        }

        private void excel导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 0;
            initInteface();

            formExcelInput frmExcelInput = new formExcelInput();
            frmExcelInput.strConn = strConn;
            frmExcelInput.strVersion = strVersion;
            if (frmExcelInput.ShowDialog(this) != DialogResult.Cancel)
                //刷新人员列表
                InitializePeopleDataGridView();
        }

        private void 增加会议ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 1;
            initInteface();

            formConference frmConference = new formConference();
            frmConference.strConn = strConn;
            frmConference.strVersion = strVersion;
            frmConference.intMode = 1; //增加模式

            frmConference.ShowDialog(this);
            //刷新会议列表
            InitializeConferenceDataGridView();
        }

        private void 修改会议ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formConference frmConference = new formConference();
            intMenu = 1;
            initInteface();

            frmConference.strConn = strConn;
            frmConference.intMode = 2; //修改模式
            frmConference.strVersion = strVersion;
            frmConference.intConferenceID = Int32.Parse(dataGridViewConference.SelectedRows[0].Cells[0].Value.ToString());

            frmConference.ShowDialog(this);

            //刷新会议列表
            InitializeConferenceDataGridView();
        }

        private void DelConferenceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((dataGridViewConference.SelectedRows.Count == 0) || (dataGridViewConference.SelectedRows[0].Cells[0].Value == null))
            {
                MessageBox.Show("请选择要删除的会议！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("是否要删除的选定的会议信息？这个操作无法恢复！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            //删除选定的会议
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
                    sqlComm.CommandText = "DELETE FROM 会议表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 分组表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 会议消费表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 参会人员表 WHERE (会议ID = " + intCID.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                    sqlComm.CommandText = "DELETE FROM 签到表 WHERE (会议ID = " + intCID.ToString() + ")";
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

        private void 消费定义ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void 会议消费ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void 会议签到ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            //刷新会议列表
            //InitializeConferenceDataGridView();
        }

        private void 忘卡补签ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 2;
            initInteface();
            formreSign frmreSign = new formreSign();

            frmreSign.strConn = strConn;
            frmreSign.intConferenceID = Int32.Parse(dataGridViewSign.SelectedRows[0].Cells[0].Value.ToString());

            frmreSign.ShowDialog(this);
        }

        private void 离线签到ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void 统计全部ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void 统计输出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void 数据库设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();

            formDatabaseSet frmDatabaseSet = new formDatabaseSet();
            frmDatabaseSet.intMode = 0;//测试模式

            if (frmDatabaseSet.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                strConn = frmDatabaseSet.strConn;
                if (strConn == "") //连接错误
                    return;

                //初始化窗口
                sqlConn.ConnectionString = strConn;
                sqlComm.Connection = sqlConn;
                sqlDA.SelectCommand = sqlComm;

                //人员管理
                InitializePeopleDataGridView();
                //会议管理
                InitializeConferenceDataGridView();
                //签到管理
                InitializedataGridViewSign();
                //消费管理
                InitializedataGridViewConsume();
                //签到统计
                InitializedataGridViewCount();
            }
        }

        private void 智能卡设置ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void 系统参数设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();

            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formSystemSet frmSystemSet = new formSystemSet();
            frmSystemSet.strVersion = strVersion;
            frmSystemSet.strConn = strConn;
            frmSystemSet.ShowDialog(this);
        }

        private void 数据管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();
            
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formDataMange frmDataMange = new formDataMange();
            frmDataMange.strConn = strConn;
            frmDataMange.ShowDialog(this);
        }

        private void 创建数据库DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            intMenu = 5;
            initInteface();

            formDatabaseSet frmDatabaseSet1 = new formDatabaseSet();
            frmDatabaseSet1.intMode = 1;//创建模式

            frmDatabaseSet1.ShowDialog(this);
            strConn = frmDatabaseSet1.strConn;
            if (strConn == "") return;

            //初始化窗口
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            //人员管理
            InitializePeopleDataGridView();
            //会议管理
            InitializeConferenceDataGridView();
            //签到管理
            InitializedataGridViewSign();
            //消费管理
            InitializedataGridViewConsume();
            //签到统计
            InitializedataGridViewCount();

        }

        private void btnSearchPeople_Click(object sender, EventArgs e)
        {
            formPeopleSearch frmPeopleSelect = new formPeopleSearch();
            frmPeopleSelect.strConn = strConn; //增加模式
            frmPeopleSelect.intCardLength = intCardLength;
            frmPeopleSelect.intCard = intCard;
            frmPeopleSelect.intComNumber = intComNumber;
            frmPeopleSelect.lComBand = lComBand;
            
            frmPeopleSelect.ShowDialog();

            if (frmPeopleSelect.intReturn == 0) return;

            //选择全部
            if (frmPeopleSelect.intReturn == -1) 
                strPSelect = "SELECT 人员ID, 姓名, 性别, 所属单位, 职称, 职务, 通讯地址, 人员类别, 主卡号 FROM 人员表";


            if (frmPeopleSelect.intReturn == 1)
            {
                bool iPosition = true;
                strPSelect = "SELECT 人员ID, 姓名, 性别, 所属单位, 职称, 职务, 通讯地址, 人员类别, 主卡号 FROM 人员表 ";
                if (frmPeopleSelect.textBoxUsername.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (姓名 LIKE N'%"+frmPeopleSelect.textBoxUsername.Text.Trim()+"%')";
                        iPosition=false;
                    }
                    else
                        strPSelect = strPSelect + " AND (姓名 LIKE N'%" + frmPeopleSelect.textBoxUsername.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.comboBoxGender.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (性别 LIKE N'%" + frmPeopleSelect.comboBoxGender.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (性别 LIKE N'%" + frmPeopleSelect.comboBoxGender.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxCompany.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (所属单位 LIKE N'%" + frmPeopleSelect.textBoxCompany.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (所属单位 LIKE N'%" + frmPeopleSelect.textBoxCompany.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxPPost.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (职称 LIKE N'%" + frmPeopleSelect.textBoxPPost.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (职称 LIKE N'%" + frmPeopleSelect.textBoxPPost.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxPost.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (职务 LIKE N'%" + frmPeopleSelect.textBoxPost.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (职务 LIKE N'%" + frmPeopleSelect.textBoxPost.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.textBoxAddress.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (通讯地址 LIKE N'%" + frmPeopleSelect.textBoxAddress.Text.Trim() + "%')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (通讯地址 LIKE N'%" + frmPeopleSelect.textBoxAddress.Text.Trim() + "%')";
                }

                if (frmPeopleSelect.comboBoxGroup.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (人员类别 = N'" + frmPeopleSelect.comboBoxGroup.Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (人员类别 = N'%" + frmPeopleSelect.comboBoxGroup.Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxCard.Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (主卡号 = N'" + frmPeopleSelect.textBoxCard.Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (主卡号 = N'%" + frmPeopleSelect.textBoxCard.Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[0].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (附加属性一 = N'" + frmPeopleSelect.textBoxAddi[0].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (附加属性一 = N'%" + frmPeopleSelect.textBoxAddi[0].Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[1].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (附加属性二 = N'" + frmPeopleSelect.textBoxAddi[1].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (附加属性二 = N'%" + frmPeopleSelect.textBoxAddi[1].Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[2].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (附加属性三 = N'" + frmPeopleSelect.textBoxAddi[2].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (附加属性三 = N'%" + frmPeopleSelect.textBoxAddi[2].Text.Trim() + "')";
                }

                if (frmPeopleSelect.textBoxAddi[3].Text.Trim() != "")
                {
                    if (iPosition)
                    {
                        strPSelect = strPSelect + " WHERE (附加属性四 = N'" + frmPeopleSelect.textBoxAddi[3].Text.Trim() + "')";
                        iPosition = false;
                    }
                    else
                        strPSelect = strPSelect + " AND (附加属性四 = N'%" + frmPeopleSelect.textBoxAddi[3].Text.Trim() + "')";
                }



            }

            InitializePeopleDataGridView();
        }

        private void btnShowAll_Click(object sender, EventArgs e)
        {
            strPSelect = "SELECT 人员ID, 姓名, 性别, 所属单位, 职称, 职务, 通讯地址, 人员类别, 主卡号 FROM 人员表";
            InitializePeopleDataGridView();
        }

        private void 人员ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            formPeopleAdd frmPeopleAdd = new formPeopleAdd(1);

            frmPeopleAdd.strConn = strConn; //增加模式
            frmPeopleAdd.strVersion = strVersion;
            frmPeopleAdd.intCardLength = intCardLength;
            frmPeopleAdd.intCard = intCard;
            frmPeopleAdd.intComNumber = intComNumber;
            frmPeopleAdd.lComBand = lComBand;
            frmPeopleAdd.bAddCon = true;


            frmPeopleAdd.ShowDialog(this);

            //刷新人员列表
            strPSelect = "SELECT 人员ID, 姓名, 性别, 所属单位, 职称, 职务, 通讯地址, 人员类别, 主卡号 FROM 人员表";
            InitializePeopleDataGridView();

        }

        private void 人员查询ToolStripMenuItem_Click(object sender, EventArgs e)
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
            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
                strSerial = dSet.Tables["软件注册"].Rows[0][0].ToString();
                dSet.Tables.Clear();

                chrislock lockit=new chrislock();

                string strData = ""; //原始码
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
                    MessageBox.Show("软件尚未进行注册？", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    boolSign = false;
                }
                else
                {
                    strVersion = strArray[2];
                    if (strArray[1] == strData || strArray[1] == "CHRISTIE(CHENYI)")
                    {
                        if (strArray[0] != "CHRISTIE(CHENYI)") //非永久码
                        {
                            try
                            {
                                System.DateTime dtRegister=System.DateTime.Parse(strArray[0]);
                                if(dtRegister < System.DateTime.Now)
                                {
                                    MessageBox.Show("软件注册信息已过期？", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    boolSign = false;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("软件注册信息不正确？", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                boolSign = false;
                            }
                            
                        }
                    }
                    else //认证码不正确
                    {
                        MessageBox.Show("软件注册信息不正确？", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        boolSign = false;
                    }
                }

            }
            else  //没有注册
            {
                strSerial = "";
                MessageBox.Show("软件尚未进行注册？", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                boolSign = false;
            }

            if (boolSign)
            {
                //
                listViewMain.Enabled = true;
                管理MToolStripMenuItem.Enabled = true;
                签到ToolStripMenuItem.Enabled = true;
                统计ToolStripMenuItem.Enabled = true;
                设置ToolStripMenuItem.Enabled = true;
            }
            else
            {
                listViewMain.Enabled = false;
                管理MToolStripMenuItem.Enabled = false;
                签到ToolStripMenuItem.Enabled = false;
                统计ToolStripMenuItem.Enabled = false;
                设置ToolStripMenuItem.Enabled = false;
            }
             */ 

            return boolSign;
        }

        private void 软件注册RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formSoftSign frmSoftSign = new formSoftSign();
            frmSoftSign.ShowDialog();
          
            if (frmSoftSign.dlResult != DialogResult.Cancel)
            {
                MessageBox.Show("软件注册完毕,请重新运行会议管理系统。", "注册", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Dispose();
            }
            
        }

        private void 附加属性管理TToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            intMenu = 0;
            initInteface();

            formPAddiMange frmPAddiMange = new formPAddiMange();
            frmPAddiMange.strConn = strConn;
            if(frmPAddiMange.ShowDialog()!=DialogResult.Cancel)
                InitializePeopleDataGridView();

        }

        private void 单位管理DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            FormDW frmDW = new FormDW();

            frmDW.strConn = strConn;
            frmDW.strVersion = strVersion;

            frmDW.ShowDialog(this);
        }

        private void 培训ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formTrainManage frmTrainManage = new formTrainManage();

            frmTrainManage.strConn = strConn;
            frmTrainManage.intMod = 0;
            frmTrainManage.strVersion = strVersion;

            frmTrainManage.ShowDialog(this);
        }

        private void 培训签到SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void 培训统计CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formTrainManage frmTrainManage = new formTrainManage();

            frmTrainManage.strConn = strConn;
            frmTrainManage.intMod = 2;
            frmTrainManage.strVersion = strVersion;

            frmTrainManage.ShowDialog(this);
        }

        private void 帮助目录CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, "itConference.chm");   
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            string strT = "人员管理";
            PrintDGV.Print_DataGridView(dataGridViewPeople, strT, false);
        }

        private void 手持机签到信息ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("数据库设置不正确！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            formSCJ frmSCJ = new formSCJ();
            if (frmSCJ.ShowDialog() != DialogResult.OK)
                return;

            int i, iR;
            string sTemp = "";

            //删除临时文件

            string dFileName = Directory.GetCurrentDirectory() + "\\ittmp.txt";
            if (File.Exists(dFileName))
            {
                File.Delete(dFileName);
            }

            cCom.sComPort = frmSCJ.comboBoxCOM.Text;
            iR = cCom.UpFile(Directory.GetCurrentDirectory());

            if (iR != 0)
            {
                MessageBox.Show("从手持机下载数据失败");
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