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
                MessageBox.Show("�������������!", "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (boolRead) //����
            {
                boolRead = false; btnRead.Text = "ǩ��"; timerS.Stop();

                saveFileDialogSign.FileName = textBoxConference.Text.Trim();
                if (saveFileDialogSign.ShowDialog() == DialogResult.OK)
                {
                    dSet.WriteXml(saveFileDialogSign.FileName);
                }
            }
            else //��ʼ
            {
                boolRead = true; btnRead.Text = "����";
                dSet.Tables["���߻�����Ϣ"].Rows[0][0] = textBoxConference.Text;
                dSet.Tables["ǩ����"].Rows.Clear();


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

            if (checkSoftSerial()) //����ע��
            {

            }
            else //û��ע��
            {
                MessageBox.Show("δ�����������������������", "ע�����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                strVersion = "3";
                //this.Close();
                //return;
            }

            //�汾����
            switch (strVersion)
            {
                case "1": //������
                case "2": //��׼��
                    MessageBox.Show("�ð汾��֧�ִ���ܣ��빺��߼��汾��", "�汾��Ϣ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
                    return;
                case "3": //������
                    break;
                default:
                    break;
            }

            dSet.Tables.Add("���߻�����Ϣ");

            dSet.Tables["���߻�����Ϣ"].Columns.Add("��������");
            dSet.Tables["���߻�����Ϣ"].Rows.Add("");

            dSet.Tables.Add("ǩ����");
            dSet.Tables["ǩ����"].Columns.Add("����", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("����", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("��λ", System.Type.GetType("System.String"));
            /*
            dSet.Tables["ǩ����"].Columns.Add("����", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("���", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("ְ��", System.Type.GetType("System.String"));
             */
            dSet.Tables["ǩ����"].Columns.Add("ǩ��ʱ��", System.Type.GetType("System.String"));
            dSet.Tables["ǩ����"].Columns.Add("ǩ��ʱ��", System.Type.GetType("System.String"));


            dataGridViewSign.DataSource = dSet.Tables["ǩ����"];

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
                //ǩ��
                signRow = dSet.Tables["ǩ����"].Select("���� = '" + textBoxSign.Text + "'");

                if (signRow.Length < 1) //�״�ǩ��
                {
                    strSignRow[0] = textBoxSign.Text;
                    strSignRow[1] = "";
                    strSignRow[2] = "";
                    strSignRow[3] = System.DateTime.Now.ToString();
                    strSignRow[4] = "";

                    dSet.Tables["ǩ����"].Rows.Add(strSignRow);
                    dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["ǩ����"].Rows.Count - 1;
                }
                else //ǩ��
                {
                    signRow[0][4] = System.DateTime.Now.ToString();
                }
                toolStripStatusLabelCount.Text = "ǩ������Ϊ��" + dataGridViewSign.Rows.Count.ToString()+ "��";
            }


        }
        

        private void formoffSign_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (boolRead) //δ����
            {
                if (MessageBox.Show("ǩ����Ϣ��δ���棿�Ƿ񱣴�ǩ����Ϣ?", "��Ϣ��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    boolRead = false; btnRead.Text = "ǩ��";

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
                this.errorProviderSign.SetError(this.textBoxConference, "�ֶι���");
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
                //ǩ��
                signRow = dSet.Tables["ǩ����"].Select("���� = '" + cMember.sCardserial + "'");

                if (signRow.Length < 1) //�״�ǩ��
                {
                    strSignRow[0] = cMember.sCardserial;
                    cMember.getUserName();
                    strSignRow[1] = cMember.strUserName;
                    strSignRow[2] = cMember.strUserDW;
                    strSignRow[3] = System.DateTime.Now.ToString();
                    strSignRow[4] = "";

                    dSet.Tables["ǩ����"].Rows.Add(strSignRow);
                    dataGridViewSign.FirstDisplayedScrollingRowIndex = dSet.Tables["ǩ����"].Rows.Count - 1;
                }
                else //ǩ��
                {
                    signRow[0][4] = System.DateTime.Now.ToString();
                }
                toolStripStatusLabelCount.Text = "ǩ������Ϊ��" + dataGridViewSign.Rows.Count.ToString() + "��";
                toolStripStatusLabelWarn.Text = "��Ч��";
                cMember.beep();
            }
            else
            {
                if (iCard==-100)
                    toolStripStatusLabelWarn.Text = "�˿�Ϊ��Ч��";
            }

        }

        private void buttonEXCEL_Click(object sender, EventArgs e)
        {
            if (boolRead) return;

            saveFileDialogExcel.FileName = textBoxConference.Text.Trim();
            if (saveFileDialogExcel.ShowDialog() != DialogResult.OK) return;


            //����һ��ExcelӦ�ó���ʵ��
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            if (excel == null)
            {
                MessageBox.Show("�޷�����Excel�ĵ���", "��������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            excel.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks[1];

            Microsoft.Office.Interop.Excel.Range r = excel.get_Range(excel.Cells[1, 1], excel.Cells[1, 1]);
            excel.Cells[1, 1] = "����";
            excel.Cells[1, 2] = "����";
            excel.Cells[1, 3] = "��λ";
            excel.Cells[1, 4] = "ǩ��ʱ��";
            excel.Cells[1, 5] = "ǩ��ʱ��";

            r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//ˮƽ����
            r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//��ֱ����
            for (int i = 0; i < dSet.Tables["ǩ����"].Rows.Count; i++)
            {
                for (int j = 0; j < dSet.Tables["ǩ����"].Columns.Count; j++)
                {
                    excel.Cells[i + 2, j + 1] = dSet.Tables["ǩ����"].Rows[i][j].ToString();
                }
            }
            Microsoft.Office.Interop.Excel.Range r1 = excel.get_Range(excel.Cells[2, 1], excel.Cells[dSet.Tables["ǩ����"].Rows.Count + 1, dSet.Tables["ǩ����"].Columns.Count]);
            //r1.Font.Name = "����";
            //r1.Font.Size = 10;
            r1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            r1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            r1.EntireColumn.AutoFit();//�Զ������п�
            r1.EntireRow.AutoFit();//�Զ������и�
            //�ӱ����
            r1.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            r1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            //���Ӵ�
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            //r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;


            workbook.SaveAs(saveFileDialogExcel.FileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            excel.Quit();
            MessageBox.Show("Execel�ļ������ɹ���");




        }
    }
}