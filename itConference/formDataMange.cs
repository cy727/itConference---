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

            if (File.Exists(dFileName)) //�����ļ�
            {
                dSet.ReadXml(dFileName);

                strDataBaseAddr = dSet.Tables["���ݿ���Ϣ"].Rows[0][0].ToString();
                strDataBaseUser = dSet.Tables["���ݿ���Ϣ"].Rows[0][1].ToString();
                strDataBasePass = dSet.Tables["���ݿ���Ϣ"].Rows[0][2].ToString();
                strDataBaseName = dSet.Tables["���ݿ���Ϣ"].Rows[0][3].ToString();
            }

            if (strDataBaseAddr == "")
            {
                btnBackUp.Enabled = false;
                btnRestore.Enabled = false;
            }
        }

        private void btnConClear_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�Ƿ�Ҫ����������ݿ⣿�ù��̲��ɻָ���", "���ݿ�", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) return;
            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM �����";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM �λ���Ա��";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM �������ѱ�";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM �����";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM ǩ����";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
        }

        private void btnPeopleClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�Ƿ�Ҫ�����Ա���ݿ⣿�ù��̲��ɻָ���", "���ݿ�", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) return;
            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM ��Ա��";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM �λ���Ա��";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM �����";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "DELETE FROM ǩ����";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
        }

        private void btnBackUp_Click(object sender, EventArgs e)
        {
            if (textBoxLJBK.Text == "")
            {
                MessageBox.Show("��ȷ�����ݿⱸ�ݵ�ַ", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            SQLDMO.Backup oBackup = new SQLDMO.BackupClass();
            SQLDMO.SQLServer oSQLServer = new SQLDMO.SQLServerClass();
            try
            {
                oSQLServer.LoginSecure = false;
                //�������õ�¼sql��������ip,��¼��,��¼����
                oSQLServer.Connect(strDataBaseAddr, strDataBaseUser, strDataBasePass);
                oBackup.Action = 0;
                //������������ʾ��������״̬
                SQLDMO.BackupSink_PercentCompleteEventHandler pceh = new SQLDMO.BackupSink_PercentCompleteEventHandler(Step2);
                oBackup.PercentComplete += pceh;
                //���ݿ�����:
                oBackup.Database = strDataBaseName;
                //���ݵ�·��
                oBackup.Files = textBoxLJBK.Text;
                //���ݵ��ļ���
                oBackup.BackupSetName = "BUSINESS";
                oBackup.BackupSetDescription = "���ݿⱸ��";
                oBackup.Initialize = true;
                oBackup.SQLBackup(oSQLServer);
                MessageBox.Show("���ݳɹ���", "��ʾ");
            }
            catch
            {
                MessageBox.Show("����ʧ�ܣ�", "��ʾ");
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
                MessageBox.Show("��ȷ�����ݿⱸ�ݵ�ַ", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�ָ����ݿ�Ḳ�ǵ�ǰ���ݣ����齨�������ݿ����д˲������������лָ�������", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
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
                MessageBox.Show("���ݿ�ָ��ɹ���");
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