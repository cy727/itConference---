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
    public partial class formExcelInput : Form
    {
        public string strConn;
        public string strVersion = "0";

        public formExcelInput()
        {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int i,j;
            
            for(i=0;i<listBoxInfo.SelectedIndices.Count;i++)
            {
                j=listBoxInfo.SelectedIndices[i];
                listBoxInput.Items.Add(listBoxInfo.Items[j]);
                listBoxInfo.Items.RemoveAt(j);

            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInput.SelectedIndices.Count; i++)
            {
                j = listBoxInput.SelectedIndices[i];
                if (listBoxInput.Items[j].ToString() == "姓名") continue;
                listBoxInfo.Items.Add(listBoxInput.Items[j]);
                listBoxInput.Items.RemoveAt(j);

            }
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
                if (listBoxInput.Items[j].ToString() == "姓名") continue;
                listBoxInfo.Items.Add(listBoxInput.Items[j]);
                listBoxInput.Items.RemoveAt(j);

            }
        }


        private void btnInput_Click(object sender, EventArgs e)
        {
            if (openFileDialogExcel.ShowDialog() == DialogResult.OK)
            {
                //string strOledbConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + openFileDialogExcel.FileName + "; 'Extended Properties=Excel 8.0;HDR=1;IMEX=1'";
                string strOledbConn = "provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + openFileDialogExcel.FileName + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";

                OleDbConnection oledbConn = new OleDbConnection(strOledbConn);

                oledbConn.Open();
                DataTable dt = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string tableName = dt.Rows[0][2].ToString().Trim();


                string strExcel = "";
                OleDbDataAdapter oledbDataAdapter = null;
                DataSet oledbdSet = new DataSet();
                strExcel = "select * from [" + tableName + "]";

                oledbDataAdapter = new OleDbDataAdapter(strExcel, oledbConn);
                oledbDataAdapter.Fill(oledbdSet, "People");
                oledbConn.Close();

                
                //导入数据
                System.Data.DataTable dtablePeople = oledbdSet.Tables["People"];
                int i,j;
                string strErr = "";

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
                        string strField = " ", strValue = " ";
                        for (j = 0; j < listBoxInput.Items.Count; j++)
                        {
                            strField = strField + listBoxInput.Items[j].ToString();
                            switch(listBoxInput.Items[j].ToString())
                            {
                                case "姓名":
                                case "性别":
                                case "职称":
                                case "职务":
                                case "所属单位":
                                case "科室":
                                case "通讯地址":
                                case "主卡号":
                                case "人员类别":
                                case "附加属性一":
                                case "附加属性二":
                                case "附加属性三":
                                case "附加属性四":
                                case "编号":

                                    strValue = strValue + "N'" + dtablePeople.Rows[i][j].ToString() + "' ";
                                    break;
                                case "年龄":
                                    strValue=strValue+" "+dtablePeople.Rows[i][j].ToString()+" ";
                                    break;
                                case "手机":
                                case "电话":
                                case "传真":
                                case "EMail":
                                case "邮政编码":

                                    strValue=strValue+"'"+dtablePeople.Rows[i][j].ToString()+"' ";
                                    break;
                                default:
                                    break;
                                         
                            }
                            if (j < listBoxInput.Items.Count - 1)
                            {
                                strField = strField + " , ";
                                strValue = strValue + " , ";
                            }

                        }

                        strErr = dtablePeople.Rows[i][0].ToString();
                        sqlComm.CommandText = "INSERT INTO 人员表 ( " + strField + " ) VALUES ( " + strValue + " )";
                        sqlComm.ExecuteNonQuery();

                        
                    }
                    sqlta.Commit();
                }
                catch
                {
                    
                    sqlta.Rollback();
                    sqlConn.Close();
                    MessageBox.Show("数据导入失败！失败数据："+strErr, "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                }
                MessageBox.Show("数据导入成功！", "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                sqlConn.Close();
                this.Close();


            }
            else
                return;
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j < 1) return;
            string strSelect = listBoxInput.Items[j-1].ToString();
            listBoxInput.Items[j - 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j - 1]);
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;

            int j = listBoxInput.SelectedIndices[0];
            if (j >= listBoxInput.Items.Count-1) return;
            string strSelect = listBoxInput.Items[j + 1].ToString();
            listBoxInput.Items[j + 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j + 1]);

        }

        private void formExcelInput_Load(object sender, EventArgs e)
        {
            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                    MessageBox.Show("该版本不支持Excel导入！请购买高级版本", "版本信息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "2": //标准版
                    break;
                case "3": //豪华版
                    break;
                default:
                    break;
            } 
        }
    }
}