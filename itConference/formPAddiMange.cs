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
    public partial class formPAddiMange : Form
    {
        public string strConn;

        public formPAddiMange()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDefine_Click(object sender, EventArgs e)
        {
            if (comboBoxIn.Text == "" || comboBoxOut.Text == "")
                return;
            if (openFileDialogAddi.ShowDialog() == DialogResult.OK)
            {
                string strOledbConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + openFileDialogAddi.FileName + ";" + "Extended Properties=\'Excel 8.0;HDR=No;\'";
                OleDbConnection oledbConn = new OleDbConnection(strOledbConn);

                oledbConn.Open();
                string strExcel = "";
                OleDbDataAdapter oledbDataAdapter = null;
                DataSet oledbdSet = new DataSet();
                strExcel = "select * from [sheet1$]";

                oledbDataAdapter = new OleDbDataAdapter(strExcel, oledbConn);
                oledbDataAdapter.Fill(oledbdSet, "People");
                oledbConn.Close();


                //导入数据
                System.Data.DataTable dtablePeople = oledbdSet.Tables["People"];
                int i, j;
                //string tt=dtablePeople.Rows[0][0].ToString();
                sqlConn.ConnectionString = strConn;
                sqlComm.Connection = sqlConn;
                sqlConn.Open();

                System.Data.SqlClient.SqlTransaction sqlta;
                sqlta = sqlConn.BeginTransaction();
                sqlComm.Transaction = sqlta;
                try
                {
                    for (i = 0; i < dtablePeople.Rows.Count; i++)
                    {
                        sqlComm.CommandText = "UPDATE 人员表 SET  " + comboBoxOut.Text + " = N'" + dtablePeople.Rows[i][1].ToString() + "' WHERE (" + comboBoxIn.Text + " = N'" + dtablePeople.Rows[i][0].ToString() + "')";
                        sqlComm.ExecuteNonQuery();
                    }
                    sqlta.Commit();
                }
                catch
                {

                    sqlta.Rollback();
                    sqlConn.Close();
                    MessageBox.Show("数据定义失败！", "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    
                }
                MessageBox.Show("数据定义成功！", "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                sqlConn.Close();
                this.Close();


            }
            else
                return;

        }
    }
}