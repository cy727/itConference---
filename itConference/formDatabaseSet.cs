using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;


namespace itConference
{
    public partial class formDatabaseSet : Form
    {
        public string strConn="";
        private string dFileName="";
        public int intMode = 0;

        private System.Data.DataSet dSet = new DataSet();

        public formDatabaseSet()
        {
            InitializeComponent();
            dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            
            strConn = "workstation id=CY;packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";;initial catalog="+textBoxDatabase.Text.Trim().ToLower();

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

            MessageBox.Show("数据库连接正常");
            sqlConn.Close();

            dSet.Tables["数据库信息"].Rows[0][0] = textBoxServer.Text;
            dSet.Tables["数据库信息"].Rows[0][1] = textBoxUser.Text;

            if(checkBoxRember.Checked) //记住密码
                dSet.Tables["数据库信息"].Rows[0][2] = textBoxPassword.Text;
            else
                dSet.Tables["数据库信息"].Rows[0][2] = "";

            dSet.Tables["数据库信息"].Rows[0][3] = textBoxDatabase.Text;
            dSet.WriteXml(dFileName);


            this.Close();

        }

        private void formDatabaseSet_Load(object sender, EventArgs e)
        {

            if (intMode == 0)//测试
            {
                btnTest.Visible = true;
                btnCreate.Visible = false;
                this.Text = "数据库设置";
            }
            else //创建
            {
                btnTest.Visible = false;
                btnCreate.Visible = true;
                this.Text = "创建数据库";
            }

            sqlComm.Connection = sqlConn;

            if(File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
            }
            else  //建立文件
            {
                dSet.Tables.Add("数据库信息");

                dSet.Tables["数据库信息"].Columns.Add("服务器地址", System.Type.GetType("System.String"));
                dSet.Tables["数据库信息"].Columns.Add("用户名", System.Type.GetType("System.String"));
                dSet.Tables["数据库信息"].Columns.Add("密码", System.Type.GetType("System.String"));
                dSet.Tables["数据库信息"].Columns.Add("数据库", System.Type.GetType("System.String"));

                string[]  strDRow ={ "","","",""};
                dSet.Tables["数据库信息"].Rows.Add(strDRow);
            }

            textBoxServer.Text = dSet.Tables["数据库信息"].Rows[0][0].ToString();
            textBoxUser.Text = dSet.Tables["数据库信息"].Rows[0][1].ToString();
            textBoxPassword.Text = dSet.Tables["数据库信息"].Rows[0][2].ToString();
            textBoxDatabase.Text = dSet.Tables["数据库信息"].Rows[0][3].ToString();

            if (textBoxServer.Text!="")
                strConn = "workstation id=CY;packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";;initial catalog=" + textBoxDatabase.Text.Trim().ToLower();
            
            
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";initial catalog=;Integrated Security=True";

            try
            {
                sqlConn.ConnectionString = strConn;
                sqlConn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库创建失败：" + ex.Message.ToString(), "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                strConn = "";
                return;
            }

            try

            {

                sqlComm.CommandText = "create database " + textBoxDatabase.Text.Trim().ToLower();
                sqlComm.ExecuteNonQuery();

                sqlConn.Close();

                strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";Database=" + textBoxDatabase.Text.Trim().ToLower() + ";Integrated Security=SSPI";
                sqlConn.ConnectionString = strConn;
                sqlConn.Open();

                sqlComm.CommandText = "CREATE TABLE [dbo].[人员分组表] ( [分组ID] [int] IDENTITY (1, 1) NOT NULL ,[分组名称] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[人员消费表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[消费ID] [int] NOT NULL ,[人员ID] [int] NOT NULL ,[消费时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[人员表] ([人员ID] [int] IDENTITY (1, 1) NOT NULL ,[姓名] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[性别] [nvarchar] (4) COLLATE Chinese_PRC_CI_AS NULL ,[年龄] [smallint] NULL ,[职称] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[职务] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[编号] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[所属单位] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[科室] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[手机] [char] (12) COLLATE Chinese_PRC_CI_AS NULL ,[电话] [char] (20) COLLATE Chinese_PRC_CI_AS NULL ,[传真] [char] (20) COLLATE Chinese_PRC_CI_AS NULL ,[EMail] [char] (50) COLLATE Chinese_PRC_CI_AS NULL ,[通讯地址] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[邮政编码] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[人员类别] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[主卡号] [nchar] (12) COLLATE Chinese_PRC_CI_AS NULL ,[附加属性一] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,[附加属性二] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,[附加属性三] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,[附加属性四] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[会议消费表] ([消费ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NULL ,[消费名称] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[开始时间] [smalldatetime] NULL ,	[结束时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[会议表] ([会议ID] [int] IDENTITY (1, 1) NOT NULL ,[会议主题] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[会议副题] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[会议内容] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[会场号] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[开始时间] [smalldatetime] NULL ,[结束时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[分组表] ([分组ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NULL ,[分组名称] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[参会人员表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NOT NULL ,[人员ID] [int] NOT NULL ,[分组名称] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[座位号] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[签到时间] [smalldatetime] NULL ,[签出时间] [smalldatetime] NULL, [短信] [smalldatetime] NULL ,[邮件] [smalldatetime] NULL ,[回复] [smalldatetime] NULL ,[留言] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[智能卡信息] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[智能卡信息] [int] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[智能卡表] ([主卡号] [char] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,[副卡一] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[副卡二] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[副卡三] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[副卡四] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[副卡五] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[签到表] ([SID] [int] NOT NULL ,[卡号] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[时间] [smalldatetime] NULL ,[会议ID] [int] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[照片表] ([人员ID] [int] NOT NULL ,[照片] [image] NULL ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[管理员表] ([AID] [int] IDENTITY (1, 1) NOT NULL ,[用户名] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,[密码] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[参会标记] ([参会] [bit] NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "INSERT INTO 参会标记 (参会) VALUES (0)";
                sqlComm.ExecuteNonQuery();


                sqlComm.CommandText = "CREATE TABLE [dbo].[会议新到人员表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NOT NULL ,[人员ID] [int] NOT NULL ,[分组名称] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[座位号] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[签到时间] [smalldatetime] NULL ,[签出时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[培训表] ([会议ID] [int] IDENTITY (1, 1) NOT NULL ,[会议主题] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[会议副题] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[会议内容] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[会场号] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[开始时间] [smalldatetime] NULL ,[结束时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[单位表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[单位编号] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[单位名称] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[上级单位] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[培训人员表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NOT NULL ,[参会单位] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[参会人数] [int] NULL ,[到会人数] [int] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[培训签到表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NOT NULL ,[人员ID] [int] NULL ,[代表单位] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[签到时间] [smalldatetime] NULL ,[签出时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[培训新到人员表] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[会议ID] [int] NOT NULL ,[人员ID] [int] NOT NULL ,[签到时间] [smalldatetime] NULL ,[签出时间] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[系统参数表](	[ID] [int] IDENTITY(1,1) NOT NULL,	[公司名] [nvarchar](100) COLLATE Chinese_PRC_CI_AS NULL,	[附加属性一] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL,	[附加属性二] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL,	[附加属性三] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL,	[附加属性四] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();
                MessageBox.Show("数据库创建成功" , "数据库", MessageBoxButtons.OK, MessageBoxIcon.Information);


                dSet.Tables["数据库信息"].Rows[0][0] = textBoxServer.Text;
                dSet.Tables["数据库信息"].Rows[0][1] = textBoxUser.Text;

                if (checkBoxRember.Checked) //记住密码
                    dSet.Tables["数据库信息"].Rows[0][2] = textBoxPassword.Text;
                else
                    dSet.Tables["数据库信息"].Rows[0][2] = "";

                dSet.Tables["数据库信息"].Rows[0][3] = textBoxDatabase.Text;
                dSet.WriteXml(dFileName);
                

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库创建失败：" + ex.Message.ToString(), "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                strConn = "";
            }
            finally
            {
                sqlConn.Close();
                this.Dispose();
            }

        }
    }
}