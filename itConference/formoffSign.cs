using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace itConference
{
    public partial class formoffSign : Form
    {

        private System.Data.DataSet dSet = new DataSet();
        private bool boolRead;
        private string[] strSignRow = { "", "", "", "", "" };
        public string strVersion = "0";
        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;
        private ClassMember cMember = new ClassMember();


        public formoffSign()
        {
            InitializeComponent();
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            timerS.Stop();
            this.Close();
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            toolStripStatusLabelWarn.Text = "";
            if (textBoxConference.Text == "")
            {
                MessageBox.Show("请输入会议名称!", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (boolRead) //结束
            {
                boolRead = false; btnRead.Text = "签到"; timerS.Stop();

                saveFileDialogSign.FileName = textBoxConference.Text.Trim();
                if (saveFileDialogSign.ShowDialog() == DialogResult.OK)
                {
                    dSet.WriteXml(saveFileDialogSign.FileName);
                }
            }
            else //开始
            {
                boolRead = true; btnRead.Text = "结束";
                dSet.Tables["离线会议信息"].Rows[0][0] = textBoxConference.Text;
                dSet.Tables["签到表"].Rows.Clear();


                switch (intCard)
                {
                    case 0:
                        textBoxSign.TextChanged += new EventHandler(textBoxSign_TextChanged);
                        this.textBoxSign.Focus();
                        this.textBoxSign.SelectAll();
                        break;
                    case 1:
                        cMember.iComPort = intComNumber;
                        cMember.lComBand = lComBand;
                        textBoxSign.TextChanged -= new EventHandler(textBoxSign_TextChanged);
                        timerS.Start();
                        //timerS_Tick(null, null);

                        break;
                }
            }

        }

        private void formoffSign_Load(object sender, EventArgs e)
        {

            if (checkSoftSerial()) //已有注册
            {

            }
            else //没有注册
            {
                MessageBox.Show("未发现软件狗，请插入软件狗。", "注册错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                strVersion = "3";
                //this.Close();
                //return;
            }

            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                case "2": //标准版
                    MessageBox.Show("该版本不支持此项功能！请购买高级版本。", "版本信息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "3": //豪华版
                    break;
                default:
                    break;
            }

            dSet.Tables.Add("离线会议信息");

            dSet.Tables["离线会议信息"].Columns.Add("会议名称");
            dSet.Tables["离线会议信息"].Rows.Add("");

            dSet.Tables.Add("签到表");
            dSet.Tables["签到表"].Columns.Add("卡号", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("姓名", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("单位", System.Type.GetType("System.String"));
            /*
            dSet.Tables["签到表"].Columns.Add("部门", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("编号", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("职务", System.Type.GetType("System.String"));
             */
            dSet.Tables["签到表"].Columns.Add("签到时间", System.Type.GetType("System.String"));
            dSet.Tables["签到表"].Columns.Add("签出时间", System.Type.GetType("System.String"));


            dataGridViewSign.DataSource = dSet.Tables["签到表"];

            dataGridViewSign.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewSign.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            /*
            dataGridViewSign.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridViewSign.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
             */
            dataGridViewSign.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            boolRead = false;

        }

        private bool checkSoftSerial()
        {
            bool boolSign = true;

            ClassSenseLock cSenseLock = new ClassSenseLock();



            cSenseLock.strModal = "itconference";
            cSenseLock.iStart = 0;
            cSenseLock.iLength = 16;

            int iR = cSenseLock.ReadSenseLock();

            if (iR != 0)
                boolSign = false;
            else
                strVersion = cSenseLock.sVersion;

            return boolSign;
        }

        
        private void textBoxSign_TextChanged(object sender, EventArgs e)
        {
            if (!boolRead) return;
            if (this.textBoxSign.Text.Length == intCardLength)
            {

                this.textBoxSign.SelectAll();

                System.Data.DataRow[] signRow;
                //签到
                signRow = dSet.Tables["签到表"].Select("卡号 = '" + textBoxSign.Text + "'");

                if (signRow.Length < 1) //首次签到
                {
                    strSignRow[0] = textBoxSign.Text;
                    strSignRow[1] = "";
                    strSignRow[2] = "";
                    strSignRow[3] = System.DateTime.Now.ToString();
                    strSignRow[4] = "";

                    dSet.Tables["签到表"].Rows.Add(strSignRow);
                    dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["签到表"].Rows.Count - 1;
                }
                else //签出
                {
                    signRow[0][4] = System.DateTime.Now.ToString();
                }
                toolStripStatusLabelCount.Text = "签到人数为：" + dataGridViewSign.Rows.Count.ToString()+ "人";
            }


        }
        

        private void formoffSign_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (boolRead) //未保存
            {
                if (MessageBox.Show("签到信息尚未保存？是否保存签到信息?", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    boolRead = false; btnRead.Text = "签到";

                    saveFileDialogSign.FileName = textBoxConference.Text.Trim();
                    if (saveFileDialogSign.ShowDialog() == DialogResult.OK)
                    {
                        dSet.WriteXml(saveFileDialogSign.FileName);
                    }
                }
            }
            timerS.Stop();
        }

        private void textBoxConference_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxConference.Text.Length > 100)
            {
                this.errorProviderSign.SetError(this.textBoxConference, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderSign.Clear();
            }
        }

        private void timerS_Tick(object sender, EventArgs e)
        {
            if (!boolRead) return;
            int iCard = 0;

            iCard = cMember.getCardSerial1();
            if (iCard == 0)
            {
                System.Data.DataRow[] signRow;
                //签到
                signRow = dSet.Tables["签到表"].Select("卡号 = '" + cMember.sCardserial + "'");

                if (signRow.Length < 1) //首次签到
                {
                    strSignRow[0] = cMember.sCardserial;
                    cMember.getUserName();
                    strSignRow[1] = cMember.strUserName;
                    strSignRow[2] = cMember.strUserDW;
                    strSignRow[3] = System.DateTime.Now.ToString();
                    strSignRow[4] = "";

                    dSet.Tables["签到表"].Rows.Add(strSignRow);
                    dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["签到表"].Rows.Count - 1;
                }
                else //签出
                {
                    signRow[0][4] = System.DateTime.Now.ToString();
                }
                toolStripStatusLabelCount.Text = "签到人数为：" + dataGridViewSign.Rows.Count.ToString() + "人";
                toolStripStatusLabelWarn.Text = "有效卡";
                cMember.beep();
            }
            else
            {
                if (iCard==-100)
                    toolStripStatusLabelWarn.Text = "此卡为无效卡";
            }

        }

        private void buttonEXCEL_Click(object sender, EventArgs e)
        {
            if (boolRead) return;

            saveFileDialogExcel.FileName = textBoxConference.Text.Trim();
            if (saveFileDialogExcel.ShowDialog() != DialogResult.OK) return;


            //创建一个Excel应用程序实例
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                MessageBox.Show("无法创建Excel文档！", "建立错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            excel.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks[1];

            Microsoft.Office.Interop.Excel.Range r = excel.get_Range(excel.Cells[1, 1], excel.Cells[1, 1]);
            excel.Cells[1, 1] = "卡号";
            excel.Cells[1, 2] = "姓名";
            excel.Cells[1, 3] = "单位";
            excel.Cells[1, 4] = "签到时间";
            excel.Cells[1, 5] = "签出时间";

            r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中
            r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//竖直居中
            for (int i = 0; i < dSet.Tables["签到表"].Rows.Count; i++)
            {
                for (int j = 0; j < dSet.Tables["签到表"].Columns.Count; j++)
                {
                    excel.Cells[i + 2, j + 1] = dSet.Tables["签到表"].Rows[i][j].ToString();
                }
            }
            Microsoft.Office.Interop.Excel.Range r1 = excel.get_Range(excel.Cells[2, 1], excel.Cells[dSet.Tables["签到表"].Rows.Count + 1, dSet.Tables["签到表"].Columns.Count]);
            //r1.Font.Name = "宋体";
            //r1.Font.Size = 10;
            r1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            r1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            r1.EntireColumn.AutoFit();//自动调整列宽
            r1.EntireRow.AutoFit();//自动调整行高
            //加表格线
            r1.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            r1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            //外框加粗
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;


            workbook.SaveAs(saveFileDialogExcel.FileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            excel.Quit();
            MessageBox.Show("Execel文件导出成功！");




        }
    }
}