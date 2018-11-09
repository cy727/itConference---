using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace itConference
{
    public partial class formDataMange : Form
    {
        public string strConn;

        public string strDataBaseAddr = "";
        public string strDataBaseUser = "";
        public string strDataBasePass = "";
        public string strDataBaseName = "";

        private string dFileName = "";

        private System.Data.DataSet dSet = new DataSet();
        private System.Data.SqlClient.SqlDataReader sqldr;

        public formDataMange()
        {
            InitializeComponent();
            dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
        }

        private void formDataMange_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;

            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);

                strDataBaseAddr = dSet.Tables["数据库信息"].Rows[0][0].ToString();
                strDataBaseUser = dSet.Tables["数据库信息"].Rows[0][1].ToString();
                strDataBasePass = dSet.Tables["数据库信息"].Rows[0][2].ToString();
                strDataBaseName = dSet.Tables["数据库信息"].Rows[0][3].ToString();
            }

            if (strDataBaseAddr == "")
            {
                btnBackUp.Enabled = false;
                btnRestore.Enabled = false;
            }
        }

        private void btnConClear_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("是否要清除会议数据库？该过程不可恢复！", "数据库", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) return;
            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM 会议表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 参会人员表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 会议消费表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 分组表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 签到表";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
        }

        private void btnPeopleClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否要清除人员数据库？该过程不可恢复！", "数据库", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) return;
            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM 人员表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 参会人员表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 分组表";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM 签到表";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
        }

        private void btnBackUp_Click(object sender, EventArgs e)
        {
            if (textBoxLJBK.Text == "")
            {
                MessageBox.Show("请确定数据库备份地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            SQLDMO.Backup oBackup = new SQLDMO.BackupClass();
            SQLDMO.SQLServer oSQLServer = new SQLDMO.SQLServerClass();
            try
            {
                oSQLServer.LoginSecure = false;
                //下面设置登录sql服务器的ip,登录名,登录密码
                oSQLServer.Connect(strDataBaseAddr, strDataBaseUser, strDataBasePass);
                oBackup.Action = 0;
                //下面两句是显示进度条的状态
                SQLDMO.BackupSink_PercentCompleteEventHandler pceh = new SQLDMO.BackupSink_PercentCompleteEventHandler(Step2);
                oBackup.PercentComplete += pceh;
                //数据库名称:
                oBackup.Database = strDataBaseName;
                //备份的路径
                oBackup.Files = textBoxLJBK.Text;
                //备份的文件名
                oBackup.BackupSetName = "BUSINESS";
                oBackup.BackupSetDescription = "数据库备份";
                oBackup.Initialize = true;
                oBackup.SQLBackup(oSQLServer);
                MessageBox.Show("备份成功！", "提示");
            }
            catch
            {
                MessageBox.Show("备份失败！", "提示");
            }
            finally
            {
                oSQLServer.DisConnect();
                toolStripProgressBar1.Value = 0;
            }
        }

        private void btnBKPath_Click(object sender, EventArgs e)
        {
            if (saveFileDialogData.ShowDialog() == DialogResult.OK)
            {
                textBoxLJBK.Text = saveFileDialogData.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialogData.ShowDialog() == DialogResult.OK)
            {
                textBoxLJHF.Text = openFileDialogData.FileName;
            }
        }

        private void Step2(string message, int percent)
        {
            toolStripProgressBar1.Value = percent;
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            if (textBoxLJHF.Text == "")
            {
                MessageBox.Show("请确定数据库备份地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("恢复数据库会覆盖当前数据，建议建立新数据库后进行此操作，继续进行恢复操作吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
            {
                return;
            }

            SQLDMO.Restore restore = new SQLDMO.RestoreClass();
            SQLDMO.SQLServer server = new SQLDMO.SQLServerClass();
            server.Connect(strDataBaseAddr, strDataBaseUser, strDataBasePass);

            //KILL DataBase Process
            sqlConn.ConnectionString = strConn;
            sqlConn.Open();
            sqlComm.CommandText = "use master Select spid FROM sysprocesses ,sysdatabases Where sysprocesses.dbid=sysdatabases.dbid AND sysdatabases.Name='" + strDataBaseName + "'";
            sqldr = sqlComm.ExecuteReader();
            while (sqldr.Read())
            {
                server.KillProcess(Convert.ToInt32(sqldr[0].ToString()));
            }
            sqldr.Close();
            sqlConn.Close();

            try
            {
                restore.Action = 0;
                SQLDMO.RestoreSink_PercentCompleteEventHandler pceh = new SQLDMO.RestoreSink_PercentCompleteEventHandler(Step2);
                restore.PercentComplete += pceh;
                restore.Database = strDataBaseName;
                restore.Files = @textBoxLJHF.Text;
                restore.ReplaceDatabase = true;
                restore.SQLRestore(server);
                MessageBox.Show("数据库恢复成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                server.DisConnect();
                toolStripProgressBar1.Value = 0;
            }
        }
    }
}