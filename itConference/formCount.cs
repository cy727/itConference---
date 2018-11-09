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
    public partial class formCount : Form
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

        public formCount()
        {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int i, j;

            for (i = 0; i < listBoxInfo.SelectedIndices.Count; i++)
            {
                j = listBoxInfo.SelectedIndices[i];
                listBoxInput.Items.Add(listBoxInfo.Items[j]);
                listBoxInfo.Items.RemoveAt(j);

            }
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

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j < 1) return;
            string strSelect = listBoxInput.Items[j - 1].ToString();
            listBoxInput.Items[j - 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j - 1]);

        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (listBoxInput.SelectedItems.Count < 1) return;
            int j = listBoxInput.SelectedIndices[0];
            if (j >= listBoxInput.Items.Count - 1) return;
            string strSelect = listBoxInput.Items[j + 1].ToString();
            listBoxInput.Items[j + 1] = listBoxInput.Items[j];
            listBoxInput.Items[j] = strSelect;
            listBoxInput.SelectedItems.Add(listBoxInput.Items[j + 1]);

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

        private void btnOffInput_Click(object sender, EventArgs e)
        {
            string PID="";
            if (openFileDialogoff.ShowDialog() != DialogResult.OK) return;
            DateTime dtTemp, dtTemp1 = System.DateTime.Now, dtTemp2 = System.DateTime.Now;
            string stTemp1 = "", stTemp2 = "";
            
            
            //System.Data.SqlClient.SqlDataReader sqldr;

            if (dSet.Tables.Contains("离线会议信息")) dSet.Tables.Remove("离线会议信息");
            if (dSet.Tables.Contains("签到表")) dSet.Tables.Remove("签到表");

            try
            {
                XmlReadMode xmlM = dSet.ReadXml(openFileDialogoff.FileName, XmlReadMode.InferSchema);
            }
            catch
            {
                MessageBox.Show("读取签到信息文件错误！", "数据库", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!dSet.Tables.Contains("签到表"))
            {
                MessageBox.Show("签到信息文件错误！", "数据库", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int i;

            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                for (i = 0; i < dSet.Tables["签到表"].Rows.Count; i++)
                {
                    if (dSet.Tables["签到表"].Rows[i][3].ToString() == "") //没有签到时间
                        continue;


                    //sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (人员表.姓名 = N'" + dSet.Tables["签到表"].Rows[i][1].ToString() + "') AND (人员表.所属单位 = N'" + dSet.Tables["签到表"].Rows[i][2].ToString() + "')";
                    sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0].ToString() + "') AND (主卡号 <> N'')";
                    sqldr = sqlComm.ExecuteReader();

                    if (!sqldr.HasRows) //无卡号
                    {
                        sqldr.Close();
                        continue;
                    }

                    //人员信息
                    sqldr.Read();
                    PID = sqldr.GetValue(14).ToString();
                    sqldr.Close();

                    //sqlComm.CommandText = "SELECT 参会人员表.人员ID FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')";
                    sqlComm.CommandText = "SELECT 参会人员表.人员ID, 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.人员ID = " + PID + ")";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows) //受邀人员
                    {
                        sqldr.Read();

                        stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][3].ToString(), dSet.Tables["签到表"].Rows[i][4].ToString());
                        stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][3].ToString(), dSet.Tables["签到表"].Rows[i][4].ToString());

                        sqldr.Close();
                        if (stTemp2 != "")
                        {
                            //sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = '" + dSet.Tables["签到表"].Rows[i][7].ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = '" + stTemp2 + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                        }
                        else
                        {
                            //sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                        }
                        sqlComm.ExecuteNonQuery();
                    }
                    else //新增人员
                    {
                        sqldr.Close();
                        //sqlComm.CommandText = "SELECT 会议新到人员表.人员ID FROM 会议新到人员表 INNER JOIN 人员表 ON 会议新到人员表.人员ID = 人员表.人员ID WHERE (会议新到人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')";
                        sqlComm.CommandText = "SELECT 会议新到人员表.人员ID, 会议新到人员表.签到时间, 会议新到人员表.签出时间 FROM 会议新到人员表 INNER JOIN 人员表 ON 会议新到人员表.人员ID = 人员表.人员ID WHERE (会议新到人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.人员ID = " + PID + ")";

                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows) //有记录
                        {
                            sqldr.Read();
                            stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][3].ToString(), dSet.Tables["签到表"].Rows[i][4].ToString());
                            stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][3].ToString(), dSet.Tables["签到表"].Rows[i][4].ToString());
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                //sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = '" + dSet.Tables["签到表"].Rows[i][7].ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = '" + stTemp2 + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                            }
                            else
                            {
                                //sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                            }
                            sqlComm.ExecuteNonQuery();
                        }
                        else //新增纪录
                        {
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                sqlComm.CommandText = "INSERT INTO 会议新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', '" + stTemp2 + "')";
                            }
                            else
                            {
                                sqlComm.CommandText = "INSERT INTO 会议新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', NULL)";
                            }
                            sqlComm.ExecuteNonQuery();
                        }

                    }
                }
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

            
            MessageBox.Show("签到信息导入完毕！", "数据库", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }

        private void formCount_Load(object sender, EventArgs e)
        {
            int INFLEN,jj;

            INFLEN=30;
            jj = INFLEN >> 1;

            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                    MessageBox.Show("该版本不支持此项功能！请购买高级版本。", "版本信息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "2": //标准版
                case "3": //豪华版
                    break;
                default:
                    break;
            } 
            
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlComm.CommandText = "SELECT 会议ID, 会议主题, 会议副题, 开始时间, 结束时间, 会场号 FROM 会议表";
            if (intConferenceID != 0) sqlComm.CommandText = sqlComm.CommandText + " WHERE (会议ID = " + intConferenceID.ToString() + ")";

            sqlConn.Open();
            sqlDA.Fill(dSet, "会议信息");

            sqlComm.CommandText = "SELECT 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            if (dSet.Tables.Contains("系统参数表")) dSet.Tables.Remove("系统参数表");
            sqlDA.Fill(dSet, "系统参数表");

            if (intConferenceID != 0) //单个会议统计
            {
                if (dSet.Tables["会议信息"].Rows[0][2].ToString() != "")
                    labelConference.Text = dSet.Tables["会议信息"].Rows[0][1].ToString() + "：" + dSet.Tables["会议信息"].Rows[0][2].ToString();
                else
                    labelConference.Text = dSet.Tables["会议信息"].Rows[0][1].ToString();

                checkBoxhztjxx.Enabled = false;
                
            }
            else
            {
                labelConference.Text = "全部会议信息统计";
                btnCount.Enabled = false;
                btnOffInput.Enabled = false;
                btnSCJ.Enabled = false;
            }
            sqlConn.Close();

            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCount_Click(object sender, EventArgs e)
        {
            string sTemp = "";

            //全部会议
            if (intConferenceID == 0) return;

            sTemp = "SELECT 人员表.人员ID, ";
            int i;
            for (i = 0; i < listBoxInput.Items.Count; i++)
            {
                sTemp = sTemp + "人员表." + listBoxInput.Items[i].ToString() + ", ";
            }


            //签到
            sqlComm.CommandText = sTemp + " 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (参会人员表.签到时间 IS NOT NULL) ORDER BY 签到时间 DESC";


            if (dSet.Tables.Contains("SIGN"))
            {
                dSet.Tables.Remove("SIGN");
                dataGridViewSign.DataSource = "";
            }
            sqlDA.Fill(dSet, "SIGN");

            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            dataGridViewSign.Columns[0].Visible = false;
            for (i = 1; i < listBoxInput.Items.Count; i++)
            {
                dataGridViewSign.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            dataGridViewSign.Columns[i+1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[i + 2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //缺席
            sqlComm.CommandText = sTemp + " 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (参会人员表.签到时间 IS NULL) ORDER BY 签到时间 DESC";


            if (dSet.Tables.Contains("UNSIGN"))
            {
                dSet.Tables.Remove("UNSIGN");
                dataGridViewUnSign.DataSource = "";
            }
            sqlDA.Fill(dSet, "UNSIGN");

            dataGridViewUnSign.DataSource = dSet.Tables["UNSIGN"];
            dataGridViewUnSign.Columns[0].Visible = false;
            for (i = 1; i < listBoxInput.Items.Count; i++)
            {
                dataGridViewUnSign.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            dataGridViewUnSign.Columns[i + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewUnSign.Columns[i + 2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewUnSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //新增
            sqlComm.CommandText = sTemp + " 会议新到人员表.签到时间, 会议新到人员表.签出时间 FROM 会议新到人员表 INNER JOIN 人员表 ON 会议新到人员表.人员ID = 人员表.人员ID WHERE (会议新到人员表.会议ID = " + intConferenceID.ToString() + ") AND (会议新到人员表.签到时间 IS NOT NULL) ORDER BY 签到时间 DESC";


            if (dSet.Tables.Contains("NEW"))
            {
                dSet.Tables.Remove("NEW");
                dataGridViewNew.DataSource = "";
            }
            sqlDA.Fill(dSet, "NEW");

            dataGridViewNew.DataSource = dSet.Tables["NEW"];
            dataGridViewNew.Columns[0].Visible = false;
            for (i = 1; i < listBoxInput.Items.Count; i++)
            {
                dataGridViewNew.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            dataGridViewNew.Columns[i + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewNew.Columns[i + 2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewNew.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            changestatusBar();


        }

        private void changestatusBar()
        {

            System.Data.SqlClient.SqlDataReader sqldr;
            string sText="";
            int iTemp=0;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 GROUP BY 会议ID HAVING (会议ID = " + dSet.Tables["会议信息"].Rows[0][0] + ")";
            sqldr = sqlComm.ExecuteReader();
            if(sqldr.HasRows)
            {
                sqldr.Read();
                sText = "登录人数： "+sqldr.GetValue(0).ToString()+" 人";
                iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
             }
            sqldr.Close();

            
             sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 WHERE (会议ID = " + dSet.Tables["会议信息"].Rows[0][0] +") AND (签到时间 IS NOT NULL)";
             sqldr = sqlComm.ExecuteReader();
            if(sqldr.HasRows)
            {
                sqldr.Read();
                sText = sText+",签到人数： "+sqldr.GetValue(0).ToString()+" 人,缺席人数： " + (iTemp - Int32.Parse(sqldr.GetValue(0).ToString())).ToString() + " 人,新增人数："+dataGridViewNew.Rows.Count.ToString()+"人";;
                sqldr.Close();
             }

            sqlConn.Close();
            toolStripStatusLabelCount.Text = sText;
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlDataReader sqldr;
            
            if (saveFileDialogOutput.ShowDialog() != DialogResult.OK) return;


            //创建一个Excel应用程序实例
            sqlConn.Open();
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();
            if (objExcel == null) 
            {
                sqlConn.Close();
                MessageBox.Show("无法创建Excel文档！", "建立错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            objExcel.Visible = false; 

            //创建一个Excel文件（未保存，无文件名）
            Workbooks objWorkbooks = objExcel.Workbooks;
            _Workbook objWorkbook = objWorkbooks.Add(XlWBATemplate.xlWBATWorksheet);//默认创建sheet1

            int i, j, k, iTemp=0,intCount=0;
            Microsoft.Office.Interop.Excel.Range range ;

            for (i = 0; i < dSet.Tables["会议信息"].Rows.Count; i++)
            {
                if(i>0) objWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, XlWBATemplate.xlWBATWorksheet);
                Sheets objSheets = objWorkbook.Worksheets;

                //表格名称
                //_Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(i+1);
                _Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(1);
                objWorksheet.Name = dSet.Tables["会议信息"].Rows[i][1].ToString();

                range = objWorksheet.get_Range("A1", Missing.Value);
                range[1, 1] = "签到情况报告";

                if (dSet.Tables["会议信息"].Rows[i][2].ToString()=="") //会议副题为空
                    range[3, 1] = "会议名称： " + dSet.Tables["会议信息"].Rows[i][1];
                else
                    range[3, 1] = "会议名称： " + dSet.Tables["会议信息"].Rows[i][1] + "：" + dSet.Tables["会议信息"].Rows[i][2];
                range[4, 1] = "开始时间： " + dSet.Tables["会议信息"].Rows[i][3] + " 结束时间： " + dSet.Tables["会议信息"].Rows[i][4];

                intCount = 5; //表格行计数
                sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 GROUP BY 会议ID HAVING (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ")";
                sqldr = sqlComm.ExecuteReader();

                if(sqldr.HasRows)
                {
                    sqldr.Read();
                    if (checkBoxderysl.Checked)
                    {
                        range[intCount, 1] = "登录人数： " + sqldr.GetValue(0).ToString() + " 人";
                        intCount++;
                    }
                    iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
                }
                sqldr.Close();


                sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 WHERE (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (签到时间 IS NOT NULL)";
               // sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 GROUP BY 会议ID, 签到时间 HAVING (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (签到时间 IS NOT NULL)";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    if (checkBoxqdrysl.Checked)
                    {
                        range[intCount, 1] = "签到人数： " + sqldr.GetValue(0).ToString().ToString() + " 人";
                        intCount++;
                    }
                    if (checkBoxqxrysl.Checked)
                    {
                        range[intCount, 1] = "缺席人数： " + (iTemp - Int32.Parse(sqldr.GetValue(0).ToString())).ToString() + " 人";
                        intCount++;
                    }
                    iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
                }
                sqldr.Close();

                if (checkBoxNew.Checked)
                {

                    sqlComm.CommandText = "SELECT COUNT(*) FROM 会议新到人员表 WHERE (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (签到时间 IS NOT NULL)";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        range[intCount, 1] = "新增人数： " + sqldr.GetValue(0).ToString().ToString() + " 人";
                        intCount++;
                    }
                    sqldr.Close();
                }


                sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 WHERE (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (签到时间 > CONVERT(DATETIME, '" + dSet.Tables["会议信息"].Rows[i][3] + "', 102))";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    if (checkBoxcdrysl.Checked)
                    {
                        iTemp = iTemp - Int32.Parse(sqldr.GetValue(0).ToString());
                        range[intCount, 1] = "准时到达： " + iTemp.ToString() + " 人， 迟到： " + sqldr.GetValue(0).ToString() + " 人";
                        intCount++;
                    }
                }
                sqldr.Close();

                //签到
                intCount++;

                if(checkBoxcdrylb.Checked) //显示迟到人员列表
                {
                if (dSet.Tables.Contains("签到信息")) dSet.Tables.Remove("签到信息");
                sqlComm.CommandText = "SELECT 人员表.人员ID, ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + "人员表." + listBoxInput.Items[j].ToString() + ", ";
                }
                if(checkBoxqcsj.Checked) //签出时间统计
                    sqlComm.CommandText = sqlComm.CommandText + " 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (参会人员表.签到时间 <= CONVERT(DATETIME,  '" + dSet.Tables["会议信息"].Rows[i][3] + "', 102)) ORDER BY 参会人员表.签到时间";
                else
                    sqlComm.CommandText = sqlComm.CommandText + " 参会人员表.签到时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (参会人员表.签到时间 <= CONVERT(DATETIME,  '" + dSet.Tables["会议信息"].Rows[i][3] + "', 102)) ORDER BY 参会人员表.签到时间";

                sqlDA.Fill(dSet, "签到信息");
                if (dSet.Tables["签到信息"].Rows.Count > 0)
                {
                    range[intCount, 1] = "人员签到情况：";
                    range[intCount + 1, 1] = "序号";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "附加属性一":
                                if(dSet.Tables["系统参数表"].Rows.Count<1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][0].ToString();
                                break;

                            case "附加属性二":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][1].ToString();
                                break;

                            case "附加属性三":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][2].ToString();
                                break;

                            case "附加属性四":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }

                        
                    }
                    range[intCount + 1, listBoxInput.Items.Count + 2] = "签到时间";
                    if(checkBoxqcsj.Checked) range[intCount + 1, listBoxInput.Items.Count + 3] = "签出时间";

                    for (j = 1; j < dSet.Tables["签到信息"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["签到信息"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["签到信息"].Rows[j - 1][k].ToString();
                        }
                    }
                    intCount += dSet.Tables["签到信息"].Rows.Count + 3;

                }


                //迟到

                if (dSet.Tables.Contains("签到信息")) dSet.Tables.Remove("签到信息");
                sqlComm.CommandText = "SELECT 人员表.人员ID, ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + "人员表." + listBoxInput.Items[j].ToString() + ", ";
                }
                sqlComm.CommandText = sqlComm.CommandText + " 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (参会人员表.签到时间 > CONVERT(DATETIME,  '" + dSet.Tables["会议信息"].Rows[i][3] + "', 102)) ORDER BY 参会人员表.签到时间";
                sqlDA.Fill(dSet, "签到信息");
                if (dSet.Tables["签到信息"].Rows.Count > 0)
                {
                    range[intCount, 1] = "迟到人员签到情况：";
                    range[intCount + 1, 1] = "序号";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "附加属性一":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][0].ToString();
                                break;

                            case "附加属性二":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][1].ToString();
                                break;

                            case "附加属性三":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][2].ToString();
                                break;

                            case "附加属性四":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }
                        
                        //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                    }
                    range[intCount + 1, listBoxInput.Items.Count + 2] = "签到时间";
                    range[intCount + 1, listBoxInput.Items.Count + 3] = "签出时间";

                    for (j = 1; j < dSet.Tables["签到信息"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["签到信息"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["签到信息"].Rows[j - 1][k].ToString();
                        }
                    }
                    intCount += dSet.Tables["签到信息"].Rows.Count + 3;

                }
                }
                else //不显示迟到列表
                {
                if (dSet.Tables.Contains("签到信息")) dSet.Tables.Remove("签到信息");
                sqlComm.CommandText = "SELECT 人员表.人员ID, ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + "人员表." + listBoxInput.Items[j].ToString() + ", ";
                }
                if(checkBoxqcsj.Checked) //签出时间统计
                    sqlComm.CommandText = sqlComm.CommandText + " 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (参会人员表.签到时间 IS NOT NULL) ORDER BY 参会人员表.签到时间";
                else
                    sqlComm.CommandText = sqlComm.CommandText + " 参会人员表.签到时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (参会人员表.签到时间 IS NOT NULL) ORDER BY 参会人员表.签到时间";

                sqlDA.Fill(dSet, "签到信息");
                if (dSet.Tables["签到信息"].Rows.Count > 0)
                {
                    range[intCount, 1] = "人员签到情况：";
                    range[intCount + 1, 1] = "序号";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "附加属性一":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][0].ToString();
                                break;

                            case "附加属性二":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][1].ToString();
                                break;

                            case "附加属性三":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][2].ToString();
                                break;

                            case "附加属性四":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }

                        //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                    }
                    range[intCount + 1, listBoxInput.Items.Count + 2] = "签到时间";
                    if(checkBoxqcsj.Checked) range[intCount + 1, listBoxInput.Items.Count + 3] = "签出时间";

                    for (j = 1; j < dSet.Tables["签到信息"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["签到信息"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["签到信息"].Rows[j - 1][k].ToString();
                        }
                    }
                    intCount += dSet.Tables["签到信息"].Rows.Count + 3;

                }

                }

                //缺席
                if(checkBoxqxrylb.Checked)
                {
                if (dSet.Tables.Contains("签到信息")) dSet.Tables.Remove("签到信息");
                sqlComm.CommandText = "SELECT 人员表.人员ID ";
                for (j = 0; j < listBoxInput.Items.Count; j++)
                {
                    sqlComm.CommandText = sqlComm.CommandText + ", 人员表." + listBoxInput.Items[j].ToString() ;
                }
                sqlComm.CommandText = sqlComm.CommandText + " FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (参会人员表.签到时间 IS NULL)";
                sqlDA.Fill(dSet, "签到信息");
                if (dSet.Tables["签到信息"].Rows.Count > 0)
                {
                    range[intCount, 1] = "缺席人员：";
                    range[intCount + 1, 1] = "序号";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        switch (listBoxInput.Items[j].ToString())
                        {
                            case "附加属性一":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][0].ToString();
                                break;

                            case "附加属性二":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][1].ToString();
                                break;

                            case "附加属性三":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][2].ToString();
                                break;

                            case "附加属性四":
                                if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                    range[intCount + 1, j + 2] = "";
                                else
                                    range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][3].ToString();
                                break;

                            default:
                                range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                break;
                        }
                        
                        //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                    }


                    for (j = 1; j < dSet.Tables["签到信息"].Rows.Count + 1; j++)
                    {
                        range[j + intCount + 1, 1] = j.ToString();
                        for (k = 1; k < dSet.Tables["签到信息"].Columns.Count; k++)
                        {
                            range[j + intCount + 1, k + 1] = dSet.Tables["签到信息"].Rows[j - 1][k].ToString();
                        }
                    }
                 }
                }


                //新增
                if (checkBoxNew.Checked)
                {

                    if (dSet.Tables.Contains("新增信息")) dSet.Tables.Remove("新增信息");
                    sqlComm.CommandText = "SELECT 人员表.人员ID ";
                    for (j = 0; j < listBoxInput.Items.Count; j++)
                    {
                        sqlComm.CommandText = sqlComm.CommandText + ", 人员表." + listBoxInput.Items[j].ToString();
                    }
                    sqlComm.CommandText = sqlComm.CommandText + ", 会议新到人员表.签到时间, 会议新到人员表.签出时间  FROM 会议新到人员表 INNER JOIN 人员表 ON 会议新到人员表.人员ID = 人员表.人员ID WHERE (会议新到人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (会议新到人员表.签到时间 IS NOT NULL)";
                    sqlDA.Fill(dSet, "新增信息");
                    if (dSet.Tables["新增信息"].Rows.Count > 0)
                    {
                        range[intCount, 1] = "新增人员：";
                        range[intCount + 1, 1] = "序号";
                        for (j = 0; j < listBoxInput.Items.Count; j++)
                        {
                            switch (listBoxInput.Items[j].ToString())
                            {
                                case "附加属性一":
                                    if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][0].ToString();
                                    break;

                                case "附加属性二":
                                    if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][1].ToString();
                                    break;

                                case "附加属性三":
                                    if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][2].ToString();
                                    break;

                                case "附加属性四":
                                    if (dSet.Tables["系统参数表"].Rows.Count < 1)
                                        range[intCount + 1, j + 2] = "";
                                    else
                                        range[intCount + 1, j + 2] = dSet.Tables["系统参数表"].Rows[0][3].ToString();
                                    break;

                                default:
                                    range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                                    break;
                            }

                            //range[intCount + 1, j + 2] = listBoxInput.Items[j].ToString();
                        }
                        range[intCount + 1, listBoxInput.Items.Count + 2] = "签到时间";
                        if (checkBoxqcsj.Checked) range[intCount + 1, listBoxInput.Items.Count + 3] = "签出时间";


                        for (j = 1; j < dSet.Tables["新增信息"].Rows.Count + 1; j++)
                        {
                            range[j + intCount + 1, 1] = j.ToString();
                            for (k = 1; k < dSet.Tables["新增信息"].Columns.Count; k++)
                            {
                                if (!checkBoxqcsj.Checked && k == dSet.Tables["新增信息"].Columns.Count - 1)
                                    continue;
                                range[j + intCount + 1, k + 1] = dSet.Tables["新增信息"].Rows[j - 1][k].ToString();
                            }
                        }
                    }
                }




            }

            if(intConferenceID==0) //全部会议统计
            {
                if (checkBoxhztjxx.Checked) //汇总统计信息
                {
                    objWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, XlWBATemplate.xlWBATWorksheet);
                    Sheets objSheets = objWorkbook.Worksheets;

                    //表格名称
                    _Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(1);
                    objWorksheet.Name = "全部会议汇总统计报告";

                    range = objWorksheet.get_Range("A1", Missing.Value);
                    sqlComm.CommandText = "SELECT COUNT(*) AS [COUNT], 参会人员表.会议ID, 会议表.会议主题, 会议表.会议副题 FROM 参会人员表 INNER JOIN  会议表 ON 参会人员表.会议ID = 会议表.会议ID WHERE (参会人员表.签到时间 IS NOT NULL) GROUP BY 参会人员表.会议ID, 会议表.会议主题, 会议表.会议副题 ORDER BY COUNT(*) DESC";
                    if (dSet.Tables.Contains("会议统计")) dSet.Tables.Remove("会议统计");
                    sqlDA.Fill(dSet, "会议统计");
                    intCount = 1;
                    if (dSet.Tables["会议统计"].Rows.Count > 0)
                    {
                        range[intCount, 1] = "全部会议签到统计：";
                        range[intCount + 1, 1] = "序号";
                        range[intCount + 1, 2] = "会议主题";
                        range[intCount + 1, 3] = "会议副题";
                        range[intCount + 1, 4] = "签到人数";
                        if (checkBoxNew.Checked)
                        {
                            range[intCount + 1, 5] = "新增人数";
                        }

                        for (j = 1; j < dSet.Tables["会议统计"].Rows.Count + 1; j++)
                        {
                            range[j + intCount + 1, 1] = j.ToString();
                            range[j + intCount + 1, 2] = dSet.Tables["会议统计"].Rows[j - 1][2].ToString();
                            range[j + intCount + 1, 3] = dSet.Tables["会议统计"].Rows[j - 1][3].ToString();
                            range[j + intCount + 1, 4] = dSet.Tables["会议统计"].Rows[j - 1][0].ToString();

                            if (checkBoxNew.Checked)
                            {

                                sqlComm.CommandText = "SELECT COUNT(*) FROM 会议新到人员表 WHERE (会议ID = " + dSet.Tables["会议统计"].Rows[j - 1][1].ToString() + ")";
                                sqldr = sqlComm.ExecuteReader();
                                if (sqldr.HasRows)
                                {
                                    sqldr.Read();
                                    range[j + intCount + 1, 5] = sqldr.GetValue(0).ToString();

                                }
                                else
                                {
                                    range[j + intCount + 1, 5] = "0";
                                }
                                sqldr.Close();
                            }

                        }
                    }
                    
                   
                }
            }

            //写入数据
 

            //保存文件
            objWorkbook._SaveAs(saveFileDialogOutput.FileName, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            //关闭文件
            objWorkbook.Close(false, saveFileDialogOutput.FileName, false);
            //objExcel.Visible = false;
            objExcel.Quit();
            objExcel = null;

            sqlConn.Close();
            MessageBox.Show("成功导出Excel文档：" + saveFileDialogOutput.FileName.ToString(), "导出", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

        }

        private void btnSCJ_Click(object sender, EventArgs e)
        {
            formSCJ frmSCJ=new formSCJ();
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
 

            string PID = "", sTemp1="";
            DateTime dtTemp, dtTemp1 = System.DateTime.Now, dtTemp2 = System.DateTime.Now;
            string stTemp1 = "", stTemp2 = "";
            DataRow[] dtC;
            bool bDT;
 
            if (dSet.Tables.Contains("离线会议信息")) dSet.Tables.Remove("离线会议信息");
            if (dSet.Tables.Contains("签到表")) dSet.Tables.Remove("签到表");

            dSet.Tables.Add("签到表");

            dSet.Tables["签到表"].Columns.Add("会员卡号", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("签到时间", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("签出时间", System.Type.GetType("System.String"));

            StreamReader swTemp = new StreamReader(dFileName, Encoding.Default);
            while (swTemp.Peek() >= 0)
            {
                string[] strDRow = {"", "", ""};
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

                //检查是否已有
                dtC = dSet.Tables["签到表"].Select("会员卡号 = '" + strDRow[0] + "'");
                if (dtC.Length > 0) //已有签到
                {
                    //签到时间
                    try
                    {
                        dtTemp1 = DateTime.Parse(dtC[0][1].ToString());
                    }
                    catch
                    {
                        continue;
                    }

                    //签出时间
                    bDT = true;
                    try
                    {
                        dtTemp2 = DateTime.Parse(dtC[0][2].ToString());
                    }
                    catch
                    {
                        bDT=false;
                    }

                    if (bDT) //有签出时间
                    {
                        if (dtTemp < dtTemp1)
                        {
                            dtC[0][1] = dtTemp.ToString();
                        }
                        else
                        {
                            if (dtTemp > dtTemp2)
                            dtC[0][2] = dtTemp.ToString();
                        }


                    }
                    else //没有签出时间
                    {
                        if(dtTemp<dtTemp1)
                        {
                            dtC[0][1]=dtTemp.ToString();
                            dtC[0][2]=dtTemp1.ToString();
                        }
                        else
                        {
                            dtC[0][2]=dtTemp.ToString();
                        }
                    }
                }
                else
                {
                    strDRow[1] = sTemp1;
                    dSet.Tables["签到表"].Rows.Add(strDRow);
                }
                
            }


            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {

                for (i = 0; i < dSet.Tables["签到表"].Rows.Count; i++)
                {
                    if (dSet.Tables["签到表"].Rows[i][1].ToString() == "") //没有签到时间
                        continue;


                    //sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (人员表.姓名 = N'" + dSet.Tables["签到表"].Rows[i][1].ToString() + "') AND (人员表.所属单位 = N'" + dSet.Tables["签到表"].Rows[i][2].ToString() + "')";
                    sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0].ToString() + "') AND (主卡号 <> N'')";
                    sqldr = sqlComm.ExecuteReader();

                    if (!sqldr.HasRows) //无卡号
                    {
                        sqldr.Close();
                        continue;
                    }

                    //人员信息
                    sqldr.Read();
                    PID = sqldr.GetValue(14).ToString();
                    sqldr.Close();

                    //sqlComm.CommandText = "SELECT 参会人员表.人员ID FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')";
                    sqlComm.CommandText = "SELECT 参会人员表.人员ID, 参会人员表.签到时间, 参会人员表.签出时间 FROM 参会人员表 INNER JOIN 人员表 ON 参会人员表.人员ID = 人员表.人员ID WHERE (参会人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.人员ID = " + PID + ")";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows) //受邀人员
                    {
                        sqldr.Read();

                        stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][1].ToString(), dSet.Tables["签到表"].Rows[i][2].ToString());
                        stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][1].ToString(), dSet.Tables["签到表"].Rows[i][2].ToString());

                        sqldr.Close();
                        if (stTemp2 != "")
                        {
                            //sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = '" + dSet.Tables["签到表"].Rows[i][7].ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = '" + stTemp2 + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                        }
                        else
                        {
                            //sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                            sqlComm.CommandText = "UPDATE 参会人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                        }
                        sqlComm.ExecuteNonQuery();
                    }
                    else //新增人员
                    {
                        sqldr.Close();
                        //sqlComm.CommandText = "SELECT 会议新到人员表.人员ID FROM 会议新到人员表 INNER JOIN 人员表 ON 会议新到人员表.人员ID = 人员表.人员ID WHERE (会议新到人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')";
                        sqlComm.CommandText = "SELECT 会议新到人员表.人员ID, 会议新到人员表.签到时间, 会议新到人员表.签出时间 FROM 会议新到人员表 INNER JOIN 人员表 ON 会议新到人员表.人员ID = 人员表.人员ID WHERE (会议新到人员表.会议ID = " + intConferenceID.ToString() + ") AND (人员表.人员ID = " + PID + ")";

                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows) //有记录
                        {
                            sqldr.Read();

                            stTemp1 = strMin(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][1].ToString(), dSet.Tables["签到表"].Rows[i][2].ToString());
                            stTemp2 = strMax(sqldr.GetValue(1).ToString().Trim(), sqldr.GetValue(2).ToString().Trim(), dSet.Tables["签到表"].Rows[i][1].ToString(), dSet.Tables["签到表"].Rows[i][2].ToString());
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                //sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = '" + dSet.Tables["签到表"].Rows[i][7].ToString() + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = '" + stTemp2 + "' WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                            }
                            else
                            {
                                //sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (主卡号 = N'" + dSet.Tables["签到表"].Rows[i][0] + "')))";
                                sqlComm.CommandText = "UPDATE 会议新到人员表 SET 签到时间 = '" + stTemp1 + "', 签出时间 = NULL WHERE (会议ID = " + intConferenceID.ToString() + ") AND (人员ID IN (SELECT 人员ID FROM 人员表 WHERE (人员ID = " + PID + ")))";
                            }
                            sqlComm.ExecuteNonQuery();
                        }
                        else //新增纪录
                        {
                            sqldr.Close();

                            if (stTemp2 != "")
                            {
                                sqlComm.CommandText = "INSERT INTO 会议新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', '" + stTemp2 + "')";
                            }
                            else
                            {
                                sqlComm.CommandText = "INSERT INTO 会议新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + stTemp1 + "', NULL)";
                            }
                            sqlComm.ExecuteNonQuery();
                        }

                    }
                }
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


            MessageBox.Show("签到信息导入完毕！", "数据库", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private string strMin(string s1, string s2, string s3, string s4)
        {
            if (s1.Trim() == "" && s2.Trim() == "" && s3.Trim() == "" && s4.Trim() == "")
                return "";

            string sTemp;
            DateTime dTemp1,dTemp2,dTemp;

            DateTime dNULL=DateTime.Parse("1900-1-1");
            dTemp1=dNULL;
            dTemp2=dNULL;
            dTemp=dNULL;


            try
            {
                dTemp1 = DateTime.Parse(s1);
            }
            catch
            {
                dTemp1 = dNULL;
            }

            try
            {
                dTemp = DateTime.Parse(s2);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp < dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s3);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp < dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s4);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp < dTemp1)
                        dTemp1 = dTemp;
                }
            }

            if (dTemp1 == dNULL)
                return "";
            else
                return dTemp1.ToString();

        }


        private string strMax(string s1, string s2, string s3, string s4)
        {
            int i;
            i = 0;
            if (s1.Trim() == "")
                i++;
            if (s2.Trim() == "")
                i++;
            if (s3.Trim() == "")
                i++;
            if (s4.Trim() == "")
                i++;

            if (i >= 3)
                return "";


            string sTemp;
            DateTime dTemp1, dTemp2, dTemp;
            DateTime dNULL = DateTime.Parse("1900-1-1");
            dTemp1 = dNULL;
            dTemp2 = dNULL;
            dTemp = dNULL;
            

            try
            {
                dTemp1 = DateTime.Parse(s1);
            }
            catch
            {
                dTemp1 = dNULL;
            }

            try
            {
                dTemp = DateTime.Parse(s2);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp > dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s3);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp > dTemp1)
                        dTemp1 = dTemp;
                }
            }

            try
            {
                dTemp = DateTime.Parse(s4);
            }
            catch
            {
                dTemp = dNULL;
            }

            if (dTemp != dNULL)
            {
                if (dTemp1 == dNULL)
                {
                    dTemp1 = dTemp;
                }
                else
                {
                    if (dTemp > dTemp1)
                        dTemp1 = dTemp;
                }
            }

            if (dTemp1 == dNULL)
                return "";
            else
                return dTemp1.ToString();
        }
    }
}