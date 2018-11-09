using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formAbout : Form
    {
        public string strVersion = "";

        public formAbout()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Dispose(); 
        }

        private void formAbout_Load(object sender, EventArgs e)
        {
            linkLabel1.Links[0].LinkData = "http://www.it922.com";

            switch (strVersion)
            {
                case "1":
                    labelVersion.Text = "(基础版)";
                    break;
                case "2":
                    labelVersion.Text = "(标准版)";
                    break;
                case "3":
                    labelVersion.Text = "(豪华版)";
                    break;
                default:
                    break;


            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string itWeb = e.Link.LinkData as string;
            System.Diagnostics.Process.Start(itWeb);
        }
    }
}