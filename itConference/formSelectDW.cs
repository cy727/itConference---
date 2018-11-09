using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formSelectDW : Form
    {
        public string strConn = "";
        public string strSelectText = "";

        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public int iDWNumber = 0;
        public string strDWName = "";
        public string strDWCode = "";

        public formSelectDW()
        {
            InitializeComponent();
        }

        private void FormFormSelectDW_Load(object sender, EventArgs e)
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

            TreeNode RootNode = new TreeNode("所有单位", 0, 1);
            int iTagRoot = 0;
            RootNode.Tag = iTagRoot;
            this.treeViewComm.Nodes.Add(RootNode);

            sqlConn.Open();
            sqlComm.CommandText = "SELECT ID, 单位编号, 单位名称, 上级单位 FROM 单位表";
            if (dSet.Tables.Contains("单位表")) dSet.Tables.Remove("单位表");
            sqlDA.Fill(dSet, "单位表");
            for (int i = 0; i < dSet.Tables["单位表"].Rows.Count; i++)
            {
                int iTag;
                if (dSet.Tables["单位表"].Rows[i][3].ToString() == "")
                    continue;
                strTemp = dSet.Tables["单位表"].Rows[i][3].ToString();
                //得到上级TAG
                iTemp = strTemp.LastIndexOf(',');
                if (iTemp != -1)
                    strTemp = strTemp.Substring(iTemp + 1);

                nodeTemp = FindTreeNodeByDepth(this.treeViewComm.Nodes, Int32.Parse(strTemp));
                TreeNode nT = new TreeNode(dSet.Tables["单位表"].Rows[i][2].ToString(), 0, 1);
                iTemp = Int32.Parse(dSet.Tables["单位表"].Rows[i][0].ToString());
                nT.Tag = iTemp;
                nodeTemp.Nodes.Add(nT);

            }
            sqlConn.Close();
            RootNode.ExpandAll();

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            iDWNumber = -1;
            this.Close();
        }

        private void treeViewComm_AfterSelect(object sender, TreeViewEventArgs e)
        {
            iDWNumber = (int)e.Node.Tag;
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (iDWNumber == 0)
                this.Close();

            sqlConn.Open();
            sqlComm.CommandText = "SELECT 单位编号, 单位名称, 上级单位 FROM 单位表 WHERE (ID = " + iDWNumber.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            while (sqldr.Read())
            {
                strDWCode = sqldr.GetValue(0).ToString();
                strDWName = sqldr.GetValue(1).ToString();
            }
            sqldr.Close();
            sqlConn.Close();

            this.Close();
        }

        private void treeViewComm_DoubleClick(object sender, EventArgs e)
        {
            btnSelect_Click(null, null);
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
    }
}
