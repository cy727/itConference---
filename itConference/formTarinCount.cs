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

namespace itConference
{
    public partial class formTarinCount : Form
    {

        public string strConn;
        public int intMode, intConferenceID;
        private System.Data.DataSet dSet = new DataSet();
        private System.Data.SqlClient.SqlDataReader sqldr;
        public string strVersion = "0";   

        public formTarinCount()
        {
            InitializeComponent();
        }

        private void formTarinCount_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlComm.CommandText = "SELECT 会议ID, 会议主题, 会议副题, 开始时间, 结束时间, 会场号 FROM 培训表";
            if (intConferenceID != 0) sqlComm.CommandText = sqlComm.CommandText + " WHERE (会议ID = " + intConferenceID.ToString() + ")";

            sqlConn.Open();
            sqlDA.Fill(dSet, "会议信息");

            if (dSet.Tables["会议信息"].Rows[0][2].ToString() != "")
               labelConference.Text = dSet.Tables["会议信息"].Rows[0][1].ToString() + "：" + dSet.Tables["会议信息"].Rows[0][2].ToString();
            else
               labelConference.Text = dSet.Tables["会议信息"].Rows[0][1].ToString();
            
            sqlConn.Close();
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
            string PID = "";
            int istart, iend;

            if (openFileDialogoff.ShowDialog() != DialogResult.OK) return;


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
            string strTemp = "",stt;

            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlComm.CommandText = "SELECT 培训人员表.参会单位, 培训人员表.参会人数, 培训人员表.到会人数, 单位表.ID, 单位表.单位编号, 单位表.上级单位 FROM 培训人员表 LEFT OUTER JOIN 单位表 ON 培训人员表.参会单位 = 单位表.单位名称 WHERE (培训人员表.会议ID = " + intConferenceID.ToString() + ")";
            sqlDA.Fill(dSet, "PEOPLE");



            sqlComm.CommandText = "SELECT ID, 单位名称, 上级单位, 单位编号 FROM 单位表 ORDER BY 上级单位";
            if (dSet.Tables.Contains("单位表")) dSet.Tables.Remove("单位表");
            sqlDA.Fill(dSet, "单位表");

            //单位表整理
            for (i = 0; i < dSet.Tables["单位表"].Rows.Count; i++)
            {
                strTemp = dSet.Tables["单位表"].Rows[i][2].ToString();
                istart = strTemp.IndexOf(',');

                if (istart == -1) //主目录
                    dSet.Tables["单位表"].Rows[i][3] = dSet.Tables["单位表"].Rows[i][0].ToString();
                else
                {
                    istart++;
                    iend = strTemp.IndexOf(',', istart);
                    if (iend == -1)
                        iend = strTemp.Length;
                    ;

                    dSet.Tables["单位表"].Rows[i][3] = strTemp.Substring(istart, iend - istart);
                }

            }

            //培训整理
            for (i = 0; i < dSet.Tables["PEOPLE"].Rows.Count; i++)
            {
                if (dSet.Tables["PEOPLE"].Rows[i][3].ToString() == "")
                    continue;

                DataRow[] dr = dSet.Tables["单位表"].Select("ID=" + dSet.Tables["PEOPLE"].Rows[i][3].ToString());
                if (dr.Length > 0)
                {
                    dSet.Tables["PEOPLE"].Rows[i][4] = dr[0][3].ToString();
                }

            }

            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            System.Data.SqlClient.SqlDataReader sqldr;

            try
            {
                for (i = 0; i < dSet.Tables["签到表"].Rows.Count; i++)
                {
                    if (dSet.Tables["签到表"].Rows[i][6].ToString() == "") //没有签到时间
                        continue;


                    sqlComm.CommandText = "SELECT 人员表.姓名, 人员表.性别, 人员表.年龄, 人员表.职称, 人员表.职务, 人员表.所属单位, 人员表.科室, 人员表.手机, 人员表.电话, 人员表.传真, 人员表.EMail, 人员表.通讯地址, 人员表.邮政编码, 人员表.主卡号,人员表.人员ID, 人员表.人员类别, 人员表.附加属性一, 人员表.附加属性二, 人员表.附加属性三, 人员表.附加属性四, 人员表.编号 FROM 人员表 WHERE (人员表.姓名 = N'" + dSet.Tables["签到表"].Rows[i][1].ToString() + "') AND (人员表.所属单位 = N'" + dSet.Tables["签到表"].Rows[i][2].ToString() + "')";
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

                    //签出
                    sqlComm.CommandText = "SELECT ID FROM 培训签到表 WHERE (人员ID = " + PID + ") AND (会议ID = " + intConferenceID.ToString() + ")";
                    sqldr = sqlComm.ExecuteReader();
                    if (sqldr.HasRows)
                    {
                        sqldr.Close();
                        continue;
                    }
                    else
                        sqldr.Close();

                    sqlComm.CommandText = "SELECT ID FROM 培训新到人员表 WHERE (人员ID = " + PID + ") AND (会议ID = " + intConferenceID.ToString() + ")";
                    sqldr = sqlComm.ExecuteReader();
                    if (sqldr.HasRows)
                    {
                        sqldr.Close();
                        continue;
                    }
                    else
                        sqldr.Close();

                    //本单位签到
                    sqlComm.CommandText = "SELECT ID FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + dSet.Tables["签到表"].Rows[i][2].ToString() + "') AND (参会人数 > 到会人数)";
                    sqldr = sqlComm.ExecuteReader();
                    if (sqldr.HasRows)
                    {
                        sqldr.Close();
                        stt = dSet.Tables["签到表"].Rows[i][7].ToString();
                        if (stt == "")
                            stt = "NULL";
                        else
                            stt = "'" + stt + "'";

                        sqlComm.CommandText = "INSERT INTO 培训签到表 (会议ID, 人员ID, 代表单位, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", N'" + dSet.Tables["签到表"].Rows[i][2].ToString() + "', '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', "+stt+")";
                        sqlComm.ExecuteNonQuery();


                        sqlComm.CommandText = "UPDATE 培训人员表 SET 到会人数 = 到会人数 + 1 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + dSet.Tables["签到表"].Rows[i][2].ToString() + "')";
                        sqlComm.ExecuteNonQuery();

                        sqldr.Close();
                        continue;

                    }
                    else
                        sqldr.Close();

                    //代表单位签到
                    while (true)
                    {
                        if (dSet.Tables["签到表"].Rows[i][2].ToString() == "")
                            break;

                        DataRow[] dr = dSet.Tables["单位表"].Select("单位名称='" + dSet.Tables["签到表"].Rows[i][2].ToString() + "'");
                        if (dr.Length < 1)
                            break;

                        strTemp = dr[0][3].ToString();

                        DataRow[] dr1 = dSet.Tables["PEOPLE"].Select("单位编号= '" + strTemp + "'");
                        if (dr1.Length < 1)
                            break;

                        strTemp = dr1[0][0].ToString();//代表单位

                        sqlComm.CommandText = "SELECT ID FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + strTemp + "') AND (参会人数 > 到会人数)";
                        sqldr = sqlComm.ExecuteReader();
                        if (sqldr.HasRows) //有代表单位
                        {
                            sqldr.Close();
                            stt = dSet.Tables["签到表"].Rows[i][7].ToString();
                            if (stt == "")
                                stt = "NULL";
                            else
                                stt = "'" + stt + "'";

                            sqlComm.CommandText = "INSERT INTO 培训签到表 (会议ID, 人员ID, 代表单位, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", N'" + strTemp + "', '" + dSet.Tables["签到表"].Rows[i][6].ToString() + "', "+stt+")";
                            sqlComm.ExecuteNonQuery();


                            sqlComm.CommandText = "UPDATE 培训人员表 SET 到会人数 = 到会人数 + 1 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会单位 = N'" + strTemp + "')";
                            sqlComm.ExecuteNonQuery();

                            sqldr.Close();
                            continue;

                        }
                        else //没有代表单位
                        {
                            sqldr.Close();
                            break;
                        }

                        break;

                    }

                    //新增
                    if (!sqldr.IsClosed) sqldr.Close();
                    sqldr.Close();
                    stt = dSet.Tables["签到表"].Rows[i][7].ToString();
                    if (stt == "")
                        stt = "NULL";
                    else
                        stt = "'" + stt + "'";

                    sqlComm.CommandText = "INSERT INTO 培训新到人员表 (会议ID, 人员ID, 签到时间, 签出时间) VALUES (" + intConferenceID.ToString() + ", " + PID + ", '" + System.DateTime.Now.ToString() + "', "+stt+")";
                    sqlComm.ExecuteNonQuery();



                    
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
            sqlComm.CommandText = sTemp + " 培训签到表.代表单位, 培训签到表.签到时间, 培训签到表.签出时间 FROM 培训签到表 LEFT OUTER JOIN  人员表 ON 培训签到表.人员ID = 人员表.人员ID WHERE (培训签到表.会议ID = " + intConferenceID.ToString() + ") ORDER BY 签到时间 DESC";


            if (dSet.Tables.Contains("SIGN"))
            {
                dSet.Tables.Remove("SIGN");
                dataGridViewSign.DataSource = "";
            }
            sqlDA.Fill(dSet, "SIGN");

            dataGridViewSign.DataSource = dSet.Tables["SIGN"];
            dataGridViewSign.Columns[0].Visible = false;
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //缺席
            sqlComm.CommandText = "SELECT 参会单位, 参会人数 , 到会人数, 参会人数 - 到会人数 AS 缺席人数 FROM 培训人员表 WHERE (会议ID = " + intConferenceID.ToString() + ") AND (参会人数 - 到会人数 <> 0)";


            if (dSet.Tables.Contains("UNSIGN"))
            {
                dSet.Tables.Remove("UNSIGN");
                dataGridViewUnSign.DataSource = "";
            }
            sqlDA.Fill(dSet, "UNSIGN");

            dataGridViewUnSign.DataSource = dSet.Tables["UNSIGN"];
            dataGridViewUnSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //新增
            sqlComm.CommandText = sTemp + " 培训新到人员表.签到时间, 培训新到人员表.签出时间 FROM 培训新到人员表 INNER JOIN 人员表 ON 培训新到人员表.人员ID = 人员表.人员ID WHERE (培训新到人员表.会议ID = " + intConferenceID.ToString() + ") ORDER BY 签到时间 DESC";


            if (dSet.Tables.Contains("NEW"))
            {
                dSet.Tables.Remove("NEW");
                dataGridViewNew.DataSource = "";
            }
            sqlDA.Fill(dSet, "NEW");

            dataGridViewNew.DataSource = dSet.Tables["NEW"];
            dataGridViewNew.Columns[0].Visible = false;
            dataGridViewNew.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            changestatusBar();


        }

        private void changestatusBar()
        {

            System.Data.SqlClient.SqlDataReader sqldr;
            string sText = "";
            int iTemp = 0;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT SUM(参会人数) AS 参会人数, 会议ID FROM 培训人员表 GROUP BY 会议ID HAVING (会议ID = " + intConferenceID.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Read();
                sText = "登录人数： " + sqldr.GetValue(0).ToString() + " 人";
                iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
                sqldr.Close();
            }


            sqlComm.CommandText = "SELECT COUNT(*) FROM 培训签到表 WHERE (会议ID = " + intConferenceID.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Read();
                sText = sText + ",签到人数： " + sqldr.GetValue(0).ToString() + " 人,缺席人数： " + (iTemp - Int32.Parse(sqldr.GetValue(0).ToString())).ToString() + " 人,新增人数：" + dataGridViewNew.Rows.Count.ToString() + "人"; ;
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
                MessageBox.Show("无法创建Excel文档！", "建立错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            objExcel.Visible = false;

            //创建一个Excel文件（未保存，无文件名）
            Workbooks objWorkbooks = objExcel.Workbooks;
            _Workbook objWorkbook = objWorkbooks.Add(XlWBATemplate.xlWBATWorksheet);//默认创建sheet1

            int i=0, j, k, iTemp = 0, intCount = 0;
            Microsoft.Office.Interop.Excel.Range range;

            Sheets objSheets = objWorkbook.Worksheets;

            //表格名称
            //_Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(i+1);
            _Worksheet objWorksheet = (_Worksheet)objSheets.get_Item(1);
            objWorksheet.Name = dSet.Tables["会议信息"].Rows[i][1].ToString();

            range = objWorksheet.get_Range("A1", Missing.Value);
            range[1, 1] = "签到情况报告";

            if (dSet.Tables["会议信息"].Rows[i][2].ToString() == "") //会议副题为空
                range[3, 1] = "会议名称： " + dSet.Tables["会议信息"].Rows[i][1];
            else
                range[3, 1] = "会议名称： " + dSet.Tables["会议信息"].Rows[i][1] + "：" + dSet.Tables["会议信息"].Rows[i][2];
            range[4, 1] = "开始时间： " + dSet.Tables["会议信息"].Rows[i][3] + " 结束时间： " + dSet.Tables["会议信息"].Rows[i][4];

            intCount = 5; //表格行计数
            sqlComm.CommandText = sqlComm.CommandText = "SELECT SUM(参会人数) AS 参会人数, 会议ID FROM 培训人员表 GROUP BY 会议ID HAVING (会议ID = " + intConferenceID.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Read();
                range[intCount, 1] = "登录人数： " + sqldr.GetValue(0).ToString() + " 人";
                intCount++;
                iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
            }
            sqldr.Close();

            sqlComm.CommandText = "SELECT COUNT(*) FROM 培训签到表 WHERE (会议ID = " + intConferenceID.ToString() + ")"; 
            // sqlComm.CommandText = "SELECT COUNT(*) FROM 参会人员表 GROUP BY 会议ID, 签到时间 HAVING (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (签到时间 IS NOT NULL)";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Read();
                range[intCount, 1] = "签到人数： " + sqldr.GetValue(0).ToString().ToString() + " 人";
                intCount++;
                 range[intCount, 1] = "缺席人数： " + (iTemp - Int32.Parse(sqldr.GetValue(0).ToString())).ToString() + " 人";
                 intCount++;
                iTemp = Int32.Parse(sqldr.GetValue(0).ToString());
            }
            sqldr.Close();

            sqlComm.CommandText = "SELECT COUNT(*) FROM 培训新到人员表 WHERE (会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (签到时间 IS NOT NULL)";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Read();
                range[intCount, 1] = "新增人数： " + sqldr.GetValue(0).ToString().ToString() + " 人";
                intCount++;
            }
            sqldr.Close();


            //签到
            intCount++;


            if (dSet.Tables.Contains("签到信息")) dSet.Tables.Remove("签到信息");
            sqlComm.CommandText = "SELECT ID, 参会单位, 参会人数, 到会人数, 参会人数 - 到会人数 AS 缺席人数 FROM 培训人员表 WHERE (会议ID = "+dSet.Tables["会议信息"].Rows[i][0]+")";

            sqlDA.Fill(dSet, "签到信息");
            if (dSet.Tables["签到信息"].Rows.Count > 0)
            {
                range[intCount, 1] = "签到情况：";
                range[intCount + 1, 1] = "序号";
                range[intCount + 1, 2] = "参会单位";
                range[intCount + 1, 3] = "参会人数";
                range[intCount + 1, 4] = "到会人数";
                range[intCount + 1, 5] = "缺席人数";


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

            //新增
            if (dSet.Tables.Contains("新增信息")) dSet.Tables.Remove("新增信息");
            sqlComm.CommandText = "SELECT 人员表.人员ID ";
            for (j = 0; j < listBoxInput.Items.Count; j++)
            {
                sqlComm.CommandText = sqlComm.CommandText + ", 人员表." + listBoxInput.Items[j].ToString();
            }
            sqlComm.CommandText = sqlComm.CommandText + ", 培训新到人员表.签到时间, 培训新到人员表.签出时间  FROM 培训新到人员表 INNER JOIN 人员表 ON 培训新到人员表.人员ID = 人员表.人员ID WHERE (培训新到人员表.会议ID = " + dSet.Tables["会议信息"].Rows[i][0] + ") AND (培训新到人员表.签到时间 IS NOT NULL)";
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


                for (j = 1; j < dSet.Tables["新增信息"].Rows.Count + 1; j++)
                {
                    range[j + intCount + 1, 1] = j.ToString();
                    for (k = 1; k < dSet.Tables["新增信息"].Columns.Count; k++)
                    {
                        range[j + intCount + 1, k + 1] = dSet.Tables["新增信息"].Rows[j - 1][k].ToString();
                    }
                }
            }

            //保存文件
            objWorkbook._SaveAs(saveFileDialogOutput.FileName, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            //关闭文件
            objWorkbook.Close(false, saveFileDialogOutput.FileName, false);
            //objExcel.Visible = false;
            objExcel.Quit();
            objExcel = null;

            sqlConn.Close();
            MessageBox.Show("成功导出Excel文档：" + saveFileDialogOutput.FileName.ToString(), "导出", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);






        }
    }
}
