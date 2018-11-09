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
    public partial class formSCJ : Form
    {
        private System.Data.DataSet dSet1 = new DataSet();

        public formSCJ()
        {
            InitializeComponent();
        }

        private void formSCJ_Load(object sender, EventArgs e)
        {
            string dFileName1 = Directory.GetCurrentDirectory() + "\\addcon.xml";

            if (File.Exists(dFileName1)) //存在文件
            {
                dSet1.ReadXml(dFileName1);
                comboBoxCOM.SelectedIndex = Int32.Parse(dSet1.Tables["附加信息"].Rows[0][0].ToString());
            }
            else
            {
                comboBoxCOM.SelectedIndex = 0;
            }
        }

        private void btnDOWN_Click(object sender, EventArgs e)
        {

        }
    }
}
