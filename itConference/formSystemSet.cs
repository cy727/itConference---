using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace itConference
{
    public partial class formSystemSet : Form
    {
        public string strConn;
        private System.Data.DataSet dSet = new DataSet();
        public string strCompanyName;
        public string strVersion = "0";

        private System.Windows.Forms.TextBox[] textBoxAddi;

        public formSystemSet()
        {
            InitializeComponent();

            this.textBoxAddi = new System.Windows.Forms.TextBox[4];
            this.textBoxAddi[0] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[1] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[2] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[3] = new System.Windows.Forms.TextBox();

            // 
            // textBoxAddi0
            // 
            this.textBoxAddi[0].Location = new System.Drawing.Point(15, 45);
            this.textBoxAddi[0].Name = "textBoxAddi0";
            this.textBoxAddi[0].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[0].TabIndex = 9;
            // 
            // textBoxAddi1
            // 
            this.textBoxAddi[1].Location = new System.Drawing.Point(15, 93);
            this.textBoxAddi[1].Name = "textBoxAddi1";
            this.textBoxAddi[1].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[1].TabIndex = 10;
            // 
            // textBoxAddi2
            // 
            this.textBoxAddi[2].Location = new System.Drawing.Point(15, 137);
            this.textBoxAddi[2].Name = "textBoxAddi2";
            this.textBoxAddi[2].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[2].TabIndex = 11;
            // 
            // textBoxAddi3
            // 
            this.textBoxAddi[3].Location = new System.Drawing.Point(15, 184);
            this.textBoxAddi[3].Name = "textBoxAddi3";
            this.textBoxAddi[3].Size = new System.Drawing.Size(179, 21);
            this.textBoxAddi[3].TabIndex = 12;
            // 

            this.groupBox3.Controls.Add(this.textBoxAddi[3]);
            this.groupBox3.Controls.Add(this.textBoxAddi[2]);
            this.groupBox3.Controls.Add(this.textBoxAddi[1]);
            this.groupBox3.Controls.Add(this.textBoxAddi[0]);

        }

        private void formSystemSet_Load(object sender, EventArgs e)
        {
            int i;
            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                case "2": //标准版
                    btnAddi.Enabled = false;
                    break;
                case "3": //豪华版
                    break;
                default:
                    break;
            } 

            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            sqlConn.Open();

            sqlComm.CommandText = "SELECT ID, 公司名, 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            sqlDA.Fill(dSet, "系统参数表");

            if (dSet.Tables["系统参数表"].Rows.Count > 0)
            {
                textBoxCompanyName.Text = dSet.Tables["系统参数表"].Rows[0][1].ToString();
                for (i = 0; i < 4; i++)
                {
                    textBoxAddi[i].Text = dSet.Tables["系统参数表"].Rows[0][i + 2].ToString();
                }
            }
            else
                textBoxCompanyName.Text = "";


            sqlComm.CommandText = "SELECT 分组ID, 分组名称 FROM 人员分组表";
            sqlDA.Fill(dSet, "人员分组表");

            if (dSet.Tables["人员分组表"].Rows.Count>0)
            {
                for(i=0;i<dSet.Tables["人员分组表"].Rows.Count;i++)
                {
                    listBoxGroup.Items.Add(dSet.Tables["人员分组表"].Rows[i][1].ToString());
                }
            }
            sqlConn.Close();
        }

        private void listBoxGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxGroup.SelectedItems.Count < 1) return;

            textBoxGroup.Text = listBoxGroup.SelectedItems[0].ToString();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if(textBoxGroup.Text.Trim()=="") return;
            listBoxGroup.Items.Add(textBoxGroup.Text.Trim());

            sqlConn.Open();
            sqlComm.CommandText = "INSERT INTO 人员分组表 (分组名称) VALUES (N'"+textBoxGroup.Text.Trim()+"')";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();

            MessageBox.Show("分组已经添加成功！");

        }

        private void btnCompanydef_Click(object sender, EventArgs e)
        {

            sqlConn.Open();
            if (dSet.Tables["系统参数表"].Rows.Count > 0)
            {
                sqlComm.CommandText = "UPDATE 系统参数表 SET 公司名 = N'" + textBoxCompanyName.Text.Trim() + "'";
            }
            else
            {
                sqlComm.CommandText = "INSERT INTO 系统参数表 (公司名) VALUES (N'"+textBoxCompanyName.Text.Trim()+"')";
            }
            sqlComm.ExecuteNonQuery();

            dSet.Tables["系统参数表"].Clear();
            sqlComm.CommandText = "SELECT ID, 公司名, 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            sqlDA.Fill(dSet, "系统参数表");

            sqlConn.Close();

            MessageBox.Show("公司名称已经修改为：‘" + textBoxCompanyName.Text.Trim()+"’");
            strCompanyName = textBoxCompanyName.Text.Trim();

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if(textBoxGroup.Text.Trim()=="") return;
            if (listBoxGroup.SelectedItems.Count < 1)
            {
                MessageBox.Show("请选择要修改的分组！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBox.Show("是否修改分组名称？该过程不可回退！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            
            sqlConn.Open();

            sqlComm.CommandText = "UPDATE 人员分组表 SET 分组名称 = N'" + textBoxGroup.Text + "' WHERE (分组名称 = N'" + listBoxGroup.SelectedItems[0].ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlComm.CommandText = "UPDATE 人员表 SET 人员类别 = N'" + textBoxGroup.Text + "' WHERE (人员类别 = N'" + listBoxGroup.SelectedItems[0].ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
            listBoxGroup.Items[listBoxGroup.SelectedIndices[0]] = textBoxGroup.Text;

            MessageBox.Show("分组名称已经修改！");

        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (listBoxGroup.SelectedItems.Count < 1)
            {
                MessageBox.Show("请选择要删除的分组！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (MessageBox.Show("是否删除分组？该过程不可回退！", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;


            sqlConn.Open();

            sqlComm.CommandText = "DELETE FROM 人员分组表 WHERE (分组名称 = N'" + listBoxGroup.SelectedItems[0] .ToString()+ "')";
            sqlComm.ExecuteNonQuery();

            sqlComm.CommandText = "UPDATE 人员表 SET 人员类别 = NULL WHERE (人员类别 = N'" + listBoxGroup.SelectedItems[0].ToString() + "')";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
            listBoxGroup.Items.Remove(listBoxGroup.SelectedItems[0]);
            listBoxGroup.SelectedItems.Clear();

            MessageBox.Show("分组名称已经删除！");
        }

        private void textBoxCompanyName_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCompanyName.Text.Length > 100)
            {
                this.errorProviders.SetError(this.textBoxCompanyName, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviders.Clear();
            }
        }

        private void textBoxGroup_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxGroup.Text.Length > 100)
            {
                this.errorProviders.SetError(this.textBoxGroup, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviders.Clear();
            }
        }

        private void btnAddi_Click(object sender, EventArgs e)
        {

            sqlConn.Open();
            if (dSet.Tables["系统参数表"].Rows.Count > 0)
            {
                sqlComm.CommandText = "UPDATE 系统参数表 SET 附加属性一 = N'" + textBoxAddi[0].Text.Trim() + "', 附加属性二 = N'" + textBoxAddi[1].Text.Trim() + "', 附加属性三 = N'" + textBoxAddi[2].Text.Trim() + "', 附加属性四 = N'"+textBoxAddi[3].Text.Trim()+"'";
            }
            else
            {
                sqlComm.CommandText = "INSERT INTO 系统参数表 (附加属性一, 附加属性二, 附加属性三, 附加属性四) VALUES (N'" + textBoxAddi[0].Text.Trim() + "', N'" + textBoxAddi[1].Text.Trim() + "', N'" + textBoxAddi[2].Text.Trim() + "', N'" + textBoxAddi[3].Text.Trim() + "')";
            }
            sqlComm.ExecuteNonQuery();

            dSet.Tables["系统参数表"].Clear();
            sqlComm.CommandText = "SELECT ID, 公司名, 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            sqlDA.Fill(dSet, "系统参数表");

            sqlConn.Close();

            MessageBox.Show("人员附加属性已经定义");

        }
    }
}