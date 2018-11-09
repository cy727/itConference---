using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class FormDW : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();
        public string strVersion = "0"; 

        public string strConn = "";

        private int iUPClass = 0;
        private string strUPClass = "0";

        private int iSelect = 0;
        
        public FormDW()
        {
            InitializeComponent();
        }

        private void FormDW_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            initCommTree();
        }

        private void initCommTree()
        {
            string strTemp;
            int iTemp;
            TreeNode nodeTemp;

            this.treeViewComm.Nodes.Clear();
            

            TreeNode RootNode = new TreeNode("全部单位", 0, 1);
            int iTagRoot = 0;
            RootNode.Tag = iTagRoot;
            this.treeViewComm.Nodes.Add(RootNode);

            sqlConn.Open();
            sqlComm.CommandText = "SELECT ID, 单位名称, 上级单位 FROM 单位表 ORDER BY 上级单位";
            if (dSet.Tables.Contains("单位表")) dSet.Tables.Remove("单位表");
            sqlDA.Fill(dSet, "单位表");

            for (int i = 0; i < dSet.Tables["单位表"].Rows.Count; i++)
            {
                int iTag;
                if (dSet.Tables["单位表"].Rows[i][2].ToString() == "")
                    continue;

                strTemp = dSet.Tables["单位表"].Rows[i][2].ToString();
                //得到上级TAG
                iTemp = strTemp.LastIndexOf(',');
                if (iTemp != -1)
                    strTemp = strTemp.Substring(iTemp + 1);

                nodeTemp = FindTreeNodeByDepth(this.treeViewComm.Nodes, Int32.Parse(strTemp));
                TreeNode nT = new TreeNode(dSet.Tables["单位表"].Rows[i][1].ToString(), 0, 1);
                iTemp = Int32.Parse(dSet.Tables["单位表"].Rows[i][0].ToString());
                nT.Tag = iTemp;
                nodeTemp.Nodes.Add(nT);

            }
            sqlConn.Close();
            RootNode.ExpandAll();

        }

        private TreeNode FindTreeNodeByDepth(TreeNodeCollection p_treeNodes, int p_i)
        {
            TreeNode treeNodeReturn = null;
            int iValue;

            foreach (TreeNode node in p_treeNodes)
            {
                //取当前节点键   
                iValue = (int)node.Tag;

                //否则根据值比   
                if (iValue == p_i)
                    treeNodeReturn = node;

                //找到即退出   
                if (treeNodeReturn != null)
                    break;
                else
                {
                    //深度优先查询   
                    if (node.Nodes.Count > 0)
                    {
                        treeNodeReturn = FindTreeNodeByDepth(node.Nodes, p_i);
                    }
                }

            }

            return treeNodeReturn;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void treeViewComm_AfterSelect(object sender, TreeViewEventArgs e)
        {
            DataRow[] drTemp;
            string strTemp = "";

            iSelect = (int)e.Node.Tag;
            if (iSelect != 0)  //
            {
                //得到上级分类
                drTemp = dSet.Tables["单位表"].Select("ID=" + iSelect.ToString());
                if (drTemp.Length < 1) //没有此类分类
                    return;

                textBoxDWMC.Text = drTemp[0][1].ToString();

                //得到上级TAG
                strTemp = drTemp[0][2].ToString();
                int iTemp = strTemp.LastIndexOf(',');
                if (iTemp != -1)
                    strTemp = strTemp.Substring(iTemp + 1);

                sqlConn.Open();
                sqlComm.CommandText = "SELECT ID, 单位名称, 上级单位 FROM 单位表 WHERE (ID = " + strTemp + ")";
                sqldr = sqlComm.ExecuteReader();

                iUPClass = 0;
                strUPClass = "0";
                textBoxSJDW.Text = "全部单位";
                while (sqldr.Read())
                {
                    iUPClass = Convert.ToInt32(sqldr.GetValue(0).ToString());
                    strUPClass = sqldr.GetValue(2).ToString() + "," + sqldr.GetValue(0).ToString();
                    textBoxSJDW.Text = sqldr.GetValue(1).ToString();
                }
                sqldr.Close();

                sqlConn.Close();


            }

        }

        private void textBoxSJDW_DoubleClick(object sender, EventArgs e)
        {
            formSelectDW frmSelectDW = new formSelectDW();
            frmSelectDW.strConn = strConn;

            frmSelectDW.ShowDialog();

            if (frmSelectDW.iDWNumber == -1)
                return;


            if (frmSelectDW.iDWNumber == 0)
            {
                strUPClass = "0";
                textBoxSJDW.Text = "";
                return;
            }

            iUPClass = frmSelectDW.iDWNumber;
            sqlConn.Open();
            sqlComm.CommandText = "SELECT ID, 单位编号, 单位名称, 上级单位  FROM 单位表 WHERE (ID = " + iUPClass.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            while (sqldr.Read())
            {
                strUPClass = sqldr.GetValue(3).ToString() + "," + sqldr.GetValue(0).ToString();
                textBoxSJDW.Text = sqldr.GetValue(2).ToString();
            }
            sqldr.Close();
            sqlConn.Close();
        }

        private void btnADD_Click(object sender, EventArgs e)
        {
            int i;
            string sTemp = "NULL";
            System.Data.SqlClient.SqlTransaction sqlta;

            if (textBoxDWMC.Text.Trim() == "")
            {
                MessageBox.Show("请输入单位名称", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            sqlConn.Open();
            //查重
            sqlComm.CommandText = "SELECT ID, 单位名称, 上级单位 FROM 单位表 WHERE (单位名称 = N'" + textBoxDWMC.Text.Trim() + "')";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Close();
                sqlConn.Close();
                MessageBox.Show("单位名称重复", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            sqldr.Close();

            
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                 sqlComm.CommandText = "INSERT INTO 单位表 (单位名称, 上级单位) VALUES (N'"+textBoxDWMC.Text.Trim()+"', N'"+strUPClass+"')";
                sqlComm.ExecuteNonQuery();

                sqlta.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库错误：" + ex.Message.ToString(), "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                return;
            }
            finally
            {
                sqlConn.Close();
            }
            MessageBox.Show("增加成功", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            initCommTree();

        }

        private void btnDEL_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlTransaction sqlta;

            if (iSelect == 0)
            {
                MessageBox.Show("请选择单位", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                sqlComm.CommandText = "DELETE FROM 单位表 WHERE (ID = " + iSelect.ToString() + ")";
                sqlComm.ExecuteNonQuery();

                sqlta.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库错误：" + ex.Message.ToString(), "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                return;
            }
            finally
            {
                sqlConn.Close();
            }
            MessageBox.Show("删除成功", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            initCommTree();
        }

        private void btnEDIT_Click(object sender, EventArgs e)
        {
            int i;
            string sTemp = "NULL";
            System.Data.SqlClient.SqlTransaction sqlta;


            if (iSelect == 0)
            {
                MessageBox.Show("请选择单位", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (textBoxDWMC.Text.Trim() == "")
            {
                MessageBox.Show("请输入单位名称", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            sqlConn.Open();
            //查重
            sqlComm.CommandText = "SELECT ID, 单位名称, 上级单位 FROM 单位表 WHERE (单位名称 = N'" + textBoxDWMC.Text.Trim() + "') AND (ID <> "+iSelect.ToString()+")";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Close();
                sqlConn.Close();
                MessageBox.Show("单位名称重复", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            sqldr.Close();

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {

                sqlComm.CommandText = "UPDATE 单位表 SET 单位名称 = N'"+textBoxDWMC.Text.Trim()+"', 上级单位 = N'"+strUPClass+"' WHERE (ID = " + iSelect.ToString() + ")";
                sqlComm.ExecuteNonQuery();

                sqlta.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库错误：" + ex.Message.ToString(), "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                return;
            }
            finally
            {
                sqlConn.Close();
            }
            MessageBox.Show("修改成功", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            initCommTree();
        }



    }
}
