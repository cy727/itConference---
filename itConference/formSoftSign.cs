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
    public partial class formSoftSign : Form
    {

        chrislock lockit=new chrislock();
        private string strpassword = "christie";
        private string dFileName = "";
        private System.Data.DataSet dSet = new DataSet();
        public DialogResult dlResult = DialogResult.Cancel;


        public formSoftSign()
        {
            InitializeComponent();
            dFileName = Directory.GetCurrentDirectory() + "\\serial.xml";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            dlResult = DialogResult.Cancel;
            this.Close();
        }


        private void formSoftSign_Load(object sender, EventArgs e)
        {
            string strData = "";

            strData = lockit.getMacAddress();
            if (strData == "")
                strData = lockit.getDiskSerial();
            if (strData == "")
            {
                strData = "CHRISTIE(CHENYI)";
            }

            textBoxSerial.Text = lockit.getLockEnData(strData, strpassword);

            if (File.Exists(dFileName)) //´æÔÚÎÄ¼þ
            {
                dSet.ReadXml(dFileName);
            }
            else  //½¨Á¢ÎÄ¼þ
            {
                dSet.Tables.Add("Èí¼þ×¢²á");

                dSet.Tables["Èí¼þ×¢²á"].Columns.Add("×¢²áÂë", System.Type.GetType("System.String"));
                string[] strDRow ={ "" };
                dSet.Tables["Èí¼þ×¢²á"].Rows.Add(strDRow);
            }

            textBoxSign.Text = dSet.Tables["Èí¼þ×¢²á"].Rows[0][0].ToString();

        }


        private void btnSign_Click(object sender, EventArgs e)
        {
            if (textBoxSign.Text.Trim() == "")
            {
                MessageBox.Show("ÇëÊäÈë×¢²áÂë£¡", "×¢²á´íÎó", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            dSet.Tables["Èí¼þ×¢²á"].Rows[0][0] = textBoxSign.Text.Trim();
            dSet.WriteXml(dFileName);

            
            dlResult = DialogResult.OK;
            this.Close();
        }
    }
}