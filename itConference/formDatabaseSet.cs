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
                MessageBox.Show("���ݿ����Ӵ����������Ա��ϵ");
                strConn = "";
                return;

            }

            MessageBox.Show("���ݿ���������");
            sqlConn.Close();

            dSet.Tables["���ݿ���Ϣ"].Rows[0][0] = textBoxServer.Text;
            dSet.Tables["���ݿ���Ϣ"].Rows[0][1] = textBoxUser.Text;

            if(checkBoxRember.Checked) //��ס����
                dSet.Tables["���ݿ���Ϣ"].Rows[0][2] = textBoxPassword.Text;
            else
                dSet.Tables["���ݿ���Ϣ"].Rows[0][2] = "";

            dSet.Tables["���ݿ���Ϣ"].Rows[0][3] = textBoxDatabase.Text;
            dSet.WriteXml(dFileName);


            this.Close();

        }

        private void formDatabaseSet_Load(object sender, EventArgs e)
        {

            if (intMode == 0)//����
            {
                btnTest.Visible = true;
                btnCreate.Visible = false;
                this.Text = "���ݿ�����";
            }
            else //����
            {
                btnTest.Visible = false;
                btnCreate.Visible = true;
                this.Text = "�������ݿ�";
            }

            sqlComm.Connection = sqlConn;

            if(File.Exists(dFileName)) //�����ļ�
            {
                dSet.ReadXml(dFileName);
            }
            else  //�����ļ�
            {
                dSet.Tables.Add("���ݿ���Ϣ");

                dSet.Tables["���ݿ���Ϣ"].Columns.Add("��������ַ", System.Type.GetType("System.String"));
                dSet.Tables["���ݿ���Ϣ"].Columns.Add("�û���", System.Type.GetType("System.String"));
                dSet.Tables["���ݿ���Ϣ"].Columns.Add("����", System.Type.GetType("System.String"));
                dSet.Tables["���ݿ���Ϣ"].Columns.Add("���ݿ�", System.Type.GetType("System.String"));

                string[]  strDRow ={ "","","",""};
                dSet.Tables["���ݿ���Ϣ"].Rows.Add(strDRow);
            }

            textBoxServer.Text = dSet.Tables["���ݿ���Ϣ"].Rows[0][0].ToString();
            textBoxUser.Text = dSet.Tables["���ݿ���Ϣ"].Rows[0][1].ToString();
            textBoxPassword.Text = dSet.Tables["���ݿ���Ϣ"].Rows[0][2].ToString();
            textBoxDatabase.Text = dSet.Tables["���ݿ���Ϣ"].Rows[0][3].ToString();

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
                MessageBox.Show("���ݿⴴ��ʧ�ܣ�" + ex.Message.ToString(), "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                sqlComm.CommandText = "CREATE TABLE [dbo].[��Ա�����] ( [����ID] [int] IDENTITY (1, 1) NOT NULL ,[��������] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��Ա���ѱ�] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NOT NULL ,[��ԱID] [int] NOT NULL ,[����ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��Ա��] ([��ԱID] [int] IDENTITY (1, 1) NOT NULL ,[����] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[�Ա�] [nvarchar] (4) COLLATE Chinese_PRC_CI_AS NULL ,[����] [smallint] NULL ,[ְ��] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[ְ��] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[���] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[������λ] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[����] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[�ֻ�] [char] (12) COLLATE Chinese_PRC_CI_AS NULL ,[�绰] [char] (20) COLLATE Chinese_PRC_CI_AS NULL ,[����] [char] (20) COLLATE Chinese_PRC_CI_AS NULL ,[EMail] [char] (50) COLLATE Chinese_PRC_CI_AS NULL ,[ͨѶ��ַ] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[��������] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[��Ա���] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[������] [nchar] (12) COLLATE Chinese_PRC_CI_AS NULL ,[��������һ] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,[�������Զ�] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,[����������] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,[����������] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[�������ѱ�] ([����ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NULL ,[��������] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[��ʼʱ��] [smalldatetime] NULL ,	[����ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[�����] ([����ID] [int] IDENTITY (1, 1) NOT NULL ,[��������] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[���鸱��] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[��������] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[�᳡��] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[��ʼʱ��] [smalldatetime] NULL ,[����ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[�����] ([����ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NULL ,[��������] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[�λ���Ա��] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NOT NULL ,[��ԱID] [int] NOT NULL ,[��������] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[��λ��] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[ǩ��ʱ��] [smalldatetime] NULL ,[ǩ��ʱ��] [smalldatetime] NULL, [����] [smalldatetime] NULL ,[�ʼ�] [smalldatetime] NULL ,[�ظ�] [smalldatetime] NULL ,[����] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[���ܿ���Ϣ] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[���ܿ���Ϣ] [int] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[���ܿ���] ([������] [char] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,[����һ] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[������] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[������] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[������] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[������] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[ǩ����] ([SID] [int] NOT NULL ,[����] [char] (10) COLLATE Chinese_PRC_CI_AS NULL ,[ʱ��] [smalldatetime] NULL ,[����ID] [int] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��Ƭ��] ([��ԱID] [int] NOT NULL ,[��Ƭ] [image] NULL ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[����Ա��] ([AID] [int] IDENTITY (1, 1) NOT NULL ,[�û���] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,[����] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[�λ���] ([�λ�] [bit] NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "INSERT INTO �λ��� (�λ�) VALUES (0)";
                sqlComm.ExecuteNonQuery();


                sqlComm.CommandText = "CREATE TABLE [dbo].[�����µ���Ա��] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NOT NULL ,[��ԱID] [int] NOT NULL ,[��������] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[��λ��] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[ǩ��ʱ��] [smalldatetime] NULL ,[ǩ��ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��ѵ��] ([����ID] [int] IDENTITY (1, 1) NOT NULL ,[��������] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[���鸱��] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[��������] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,[�᳡��] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[��ʼʱ��] [smalldatetime] NULL ,[����ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��λ��] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[��λ���] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,[��λ����] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,[�ϼ���λ] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��ѵ��Ա��] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NOT NULL ,[�λᵥλ] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[�λ�����] [int] NULL ,[��������] [int] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��ѵǩ����] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NOT NULL ,[��ԱID] [int] NULL ,[����λ] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,[ǩ��ʱ��] [smalldatetime] NULL ,[ǩ��ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[��ѵ�µ���Ա��] ([ID] [int] IDENTITY (1, 1) NOT NULL ,[����ID] [int] NOT NULL ,[��ԱID] [int] NOT NULL ,[ǩ��ʱ��] [smalldatetime] NULL ,[ǩ��ʱ��] [smalldatetime] NULL ) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[ϵͳ������](	[ID] [int] IDENTITY(1,1) NOT NULL,	[��˾��] [nvarchar](100) COLLATE Chinese_PRC_CI_AS NULL,	[��������һ] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL,	[�������Զ�] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL,	[����������] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL,	[����������] [nvarchar](50) COLLATE Chinese_PRC_CI_AS NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();
                MessageBox.Show("���ݿⴴ���ɹ�" , "���ݿ�", MessageBoxButtons.OK, MessageBoxIcon.Information);


                dSet.Tables["���ݿ���Ϣ"].Rows[0][0] = textBoxServer.Text;
                dSet.Tables["���ݿ���Ϣ"].Rows[0][1] = textBoxUser.Text;

                if (checkBoxRember.Checked) //��ס����
                    dSet.Tables["���ݿ���Ϣ"].Rows[0][2] = textBoxPassword.Text;
                else
                    dSet.Tables["���ݿ���Ϣ"].Rows[0][2] = "";

                dSet.Tables["���ݿ���Ϣ"].Rows[0][3] = textBoxDatabase.Text;
                dSet.WriteXml(dFileName);
                

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("���ݿⴴ��ʧ�ܣ�" + ex.Message.ToString(), "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
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