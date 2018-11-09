using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.IO;


namespace itConference
{
    public partial class formSCJQD : Form
    {
        public string strConn;
        public int intMode, intConferenceID;
        private System.Data.DataSet dSet = new DataSet();
        private System.Data.SqlClient.SqlDataReader sqldr;
        public string strVersion = "0";
        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;

        private int iLENGTH = 30;
        private int iCLENGTH = 10;

        private classCOM cCom = new classCOM();
        private ClassMember cMember = new ClassMember();

        public formSCJQD()
        {
            InitializeComponent();
        }

        private void formSCJQD_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlComm.CommandText = "SELECT 人员ID, 姓名, 主卡号 FROM  人员表 WHERE (主卡号 <> N'') AND (主卡号 IS NOT NULL)";

            sqlConn.Open();
            sqlDA.Fill(dSet, "人员表");
            sqlConn.Close();

            if (dSet.Tables.Contains("签到时间表")) dSet.Tables.Remove("签到时间表");
            if (dSet.Tables.Contains("签到表")) dSet.Tables.Remove("签到表");

            dSet.Tables.Add("签到表");

            dSet.Tables["签到表"].Columns.Add("人员ID", System.Type.GetType("System.Int32"));
            dSet.Tables["签到表"].Columns.Add("姓名", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("签到次数", System.Type.GetType("System.Int32"));

            dSet.Tables.Add("签到时间表");

            dSet.Tables["签到时间表"].Columns.Add("人员ID", System.Type.GetType("System.Int32"));
            dSet.Tables["签到时间表"].Columns.Add("姓名", System.Type.GetType("System.String"));
            dSet.Tables["签到时间表"].Columns.Add("签到次数", System.Type.GetType("System.Int32"));
            dSet.Tables["签到时间表"].Columns.Add("签到时间", System.Type.GetType("System.String"));


            string dFileName = Directory.GetCurrentDirectory() + "\\ittmp.txt";
            if (!File.Exists(dFileName))
            {
                this.Close();
                return;
            }

            string sTemp1 = "";
            DateTime dtTemp, dtTemp1 = System.DateTime.Now, dtTemp2 = System.DateTime.Now;
            string stTemp1 = "", stTemp2 = "";
            DataRow[] dtC;
            string sTemp = "";
            int i, j;

            StreamReader swTemp = new StreamReader(dFileName, Encoding.Default);
            while (swTemp.Peek() >= 0)
            {
                string[] strDRow = { "", "", "" };
                sTemp = swTemp.ReadLine();
                if (sTemp.Length < iLENGTH)
                    continue;

                stTemp1 = sTemp.Substring(0, iCLENGTH).Trim();

                if (intCard == 0) //ID卡，进行转换
                    strDRow[0] = cMember.IDCARDNOFROMHEX(stTemp1, intCardLength);
                else
                    strDRow[0] = stTemp1;

                if (strDRow[0] == "0000000000")
                    continue;

                sTemp1 = sTemp.Substring(iCLENGTH + 1, 19);
                try
                {
                    dtTemp = DateTime.Parse(sTemp1);
                }
                catch
                {
                    continue;
                }

                //得到人员信息
                //检查是有该人
                dtC = dSet.Tables["人员表"].Select("主卡号 = '" + strDRow[0] + "'");
                if (dtC.Length < 1)
                    continue;

                DataRow[] dtC1;
                dtC1 = dSet.Tables["签到表"].Select("人员ID = " + dtC[0][0] + "");

                if (dtC1.Length < 1) //没有该人记录
                {
                    object[] dr1 = new object[3];
                    dr1[0] = dtC[0][0];
                    dr1[1] = dtC[0][1];
                    dr1[2] = 1;
                    dSet.Tables["签到表"].Rows.Add(dr1);
                }
                else
                {
                    dtC1[0][2] = int.Parse(dtC1[0][2].ToString()) + 1;
                }


                object[] dr2 = new object[4];
                dr2[0] = dtC[0][0];
                dr2[1] = dtC[0][1];
                dr2[2] = 1;
                dr2[3] = dtTemp.ToString();
                dSet.Tables["签到时间表"].Rows.Add(dr2);

            }

            for (i = 0; i < dSet.Tables["签到表"].Rows.Count; i++)
            {
                DataRow[] dtC2;
                dtC2 = dSet.Tables["签到时间表"].Select("人员ID = " + dSet.Tables["签到表"].Rows[i][0].ToString() + "");

                for(j=0;j<dtC2.Length;j++)
                {
                    dtC2[j][2]=int.Parse(dSet.Tables["签到表"].Rows[i][2].ToString());
                }
            }

            button1_Click(null,null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i;
            if (checkBoxSJ.Checked)
            {
                DataView dt = new DataView(dSet.Tables["签到时间表"]);
                if (!checkBoxAll.Checked)
                {
                    dt.RowFilter = "姓名 LIKE '%"+textBoxXM.Text+"%'";
                }
                dataGridViewMX.DataSource = dt;
            }
            else
            {
                DataView dt = new DataView(dSet.Tables["签到表"]);
                if (!checkBoxAll.Checked)
                {
                    dt.RowFilter = "姓名 LIKE '%" + textBoxXM.Text + "%'";
                }
                dataGridViewMX.DataSource = dt;
            }

            dataGridViewMX.Columns[0].Visible = false;

            for(i=1;i<dataGridViewMX.ColumnCount;i++)
                dataGridViewMX.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            

        }

        private void buttonPrn_Click(object sender, EventArgs e)
        {
            string strT = "签到管理";
            PrintDGV.Print_DataGridView(dataGridViewMX, strT, true);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
