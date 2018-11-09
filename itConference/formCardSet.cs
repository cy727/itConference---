using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace itConference
{
    public partial class formCardSet : Form
    {
        public int intLength=10;
        public int intComNumber = 1, intCardLength = 10, intCard = 1;
        public int lComBand = 19200;
        private string dFileName = "";
        private System.Data.DataSet dSet = new DataSet();
        private ClassMember cMember = new ClassMember();

        private string dFileName1 = "";
        private System.Data.DataSet dSet1 = new DataSet();

        public int intComAdd1 = 1, intComAdd2 = 1;

        public formCardSet()
        {
            InitializeComponent();
            dFileName = Directory.GetCurrentDirectory() + "\\cardcon.xml";

            dFileName1 = Directory.GetCurrentDirectory() + "\\addcon.xml";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSet_Click(object sender, EventArgs e)
        {
            intLength = Int32.Parse(textBoxCardLength.Text.Trim());
            intComNumber = comboBoxCom.SelectedIndex;
            lComBand = Int32.Parse(comboBoxBand.Text);
            intCard = comboBoxCard.SelectedIndex;

            dSet.Tables["智能卡信息"].Rows[0][0] = comboBoxCard.SelectedIndex.ToString();
            dSet.Tables["智能卡信息"].Rows[0][1] = comboBoxCom.SelectedIndex.ToString();
            dSet.Tables["智能卡信息"].Rows[0][2] = textBoxCardLength.Text;
            dSet.Tables["智能卡信息"].Rows[0][3] = comboBoxBand.Text;

            dSet.WriteXml(dFileName);

            dSet1.Tables["附加信息"].Rows[0][0] = comboBoxAdd1.SelectedIndex.ToString();
            dSet1.Tables["附加信息"].Rows[0][1] = comboBoxAdd2.SelectedIndex.ToString();

            dSet1.WriteXml(dFileName1);
            this.Close();
        }

        private void formCardSet_Load(object sender, EventArgs e)
        {
            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
            }
            else  //建立文件
            {
                dSet.Tables.Add("智能卡信息");

                dSet.Tables["智能卡信息"].Columns.Add("类型", System.Type.GetType("System.String"));
                dSet.Tables["智能卡信息"].Columns.Add("端口号", System.Type.GetType("System.String"));
                dSet.Tables["智能卡信息"].Columns.Add("卡号长度", System.Type.GetType("System.String"));
                dSet.Tables["智能卡信息"].Columns.Add("通讯波特率", System.Type.GetType("System.String"));

                string[] strDRow ={ "1", "1", intLength .ToString(),"19200"};
                dSet.Tables["智能卡信息"].Rows.Add(strDRow);
            }

            if (File.Exists(dFileName1)) //存在文件
            {
                dSet1.ReadXml(dFileName1);
            }
            else  //建立文件
            {
                dSet1.Tables.Add("附加信息");

                dSet1.Tables["附加信息"].Columns.Add("端口号1", System.Type.GetType("System.String"));
                dSet1.Tables["附加信息"].Columns.Add("端口号2", System.Type.GetType("System.String"));

                string[] strDRow1 = { "1", "1"};
                dSet1.Tables["附加信息"].Rows.Add(strDRow1);
            }


            comboBoxCard.SelectedIndex = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][0].ToString());
            comboBoxCom.SelectedIndex = Int32.Parse(dSet.Tables["智能卡信息"].Rows[0][1].ToString());
            textBoxCardLength.Text=dSet.Tables["智能卡信息"].Rows[0][2].ToString();
            comboBoxBand.Text = dSet.Tables["智能卡信息"].Rows[0][3].ToString();

            comboBoxAdd1.SelectedIndex = Int32.Parse(dSet1.Tables["附加信息"].Rows[0][0].ToString());
            comboBoxAdd2.SelectedIndex = Int32.Parse(dSet1.Tables["附加信息"].Rows[0][1].ToString());


        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            if (comboBoxCard.SelectedIndex != 1)
                return;

            cMember.iComPort = comboBoxCom.SelectedIndex;
            cMember.lComBand = Int32.Parse(comboBoxBand.Text);

            cMember.beep();
            if (cMember.getCardSerial() == 0)
                MessageBox.Show("读卡成功！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                MessageBox.Show("读卡失败，请检查读卡器及其设置是否正确！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
    }
}