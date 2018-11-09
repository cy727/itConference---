using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace itConference
{
    public partial class formPeopleNow : Form
    {
        public int iMode,intPID,intCardLength=10;
        public string strConn;
        private System.Data.DataSet dSet = new DataSet();
        private bool boolDelPic=false;
        public string strVersion = "0";

        private System.Windows.Forms.Label[] labelAddi;
        private System.Windows.Forms.TextBox[] textBoxAddi;

        public formPeopleNow(int intMode)
        {
            InitializeComponent();
            // 
            // labelAddi0
            // 
            this.labelAddi = new System.Windows.Forms.Label[4];
            this.labelAddi[0] = new System.Windows.Forms.Label();
            this.labelAddi[1] = new System.Windows.Forms.Label();
            this.labelAddi[2] = new System.Windows.Forms.Label();
            this.labelAddi[3] = new System.Windows.Forms.Label();
            this.textBoxAddi = new System.Windows.Forms.TextBox[4];
            this.textBoxAddi[0] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[1] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[2] = new System.Windows.Forms.TextBox();
            this.textBoxAddi[3] = new System.Windows.Forms.TextBox();

            this.labelAddi[0].AutoSize = true;
            this.labelAddi[0].Location = new System.Drawing.Point(8, 21);
            this.labelAddi[0].Name = "labelAddi0";
            this.labelAddi[0].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[0].TabIndex = 0;
            this.labelAddi[0].Text = "附加属性一：";
            // 
            // labelAddi1
            // 
            this.labelAddi[1].AutoSize = true;
            this.labelAddi[1].Location = new System.Drawing.Point(244, 21);
            this.labelAddi[1].Name = "labelAddi1";
            this.labelAddi[1].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[1].TabIndex = 1;
            this.labelAddi[1].Text = "附加属性二：";
            // 
            // labelAddi2
            // 
            this.labelAddi[2].AutoSize = true;
            this.labelAddi[2].Location = new System.Drawing.Point(8, 48);
            this.labelAddi[2].Name = "labelAddi2";
            this.labelAddi[2].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[2].TabIndex = 2;
            this.labelAddi[2].Text = "附加属性三：";
            // 
            // labelAddi3
            // 
            this.labelAddi[3].AutoSize = true;
            this.labelAddi[3].Location = new System.Drawing.Point(244, 48);
            this.labelAddi[3].Name = "labelAddi3";
            this.labelAddi[3].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[3].TabIndex = 3;
            this.labelAddi[3].Text = "附加属性四：";
            // 
            // textBoxAddi0
            // 
            this.textBoxAddi[0].Location = new System.Drawing.Point(82, 18);
            this.textBoxAddi[0].Name = "textBoxAddi0";
            this.textBoxAddi[0].Size = new System.Drawing.Size(150, 21);
            this.textBoxAddi[0].TabIndex = 4;
            // 
            // textBoxAddi1
            // 
            this.textBoxAddi[1].Location = new System.Drawing.Point(317, 17);
            this.textBoxAddi[1].Name = "textBoxAddi1";
            this.textBoxAddi[1].Size = new System.Drawing.Size(150, 21);
            this.textBoxAddi[1].TabIndex = 5;
            // 
            // textBoxAddi2
            // 
            this.textBoxAddi[2].Location = new System.Drawing.Point(82, 45);
            this.textBoxAddi[2].Name = "textBoxAddi2";
            this.textBoxAddi[2].Size = new System.Drawing.Size(150, 21);
            this.textBoxAddi[2].TabIndex = 6;
            // 
            // textBoxAddi3
            // 
            this.textBoxAddi[3].Location = new System.Drawing.Point(317, 44);
            this.textBoxAddi[3].Name = "textBoxAddi3";
            this.textBoxAddi[3].Size = new System.Drawing.Size(150, 21);
            this.textBoxAddi[3].TabIndex = 7;

            this.groupBox3.Controls.Add(this.textBoxAddi[3]);
            this.groupBox3.Controls.Add(this.textBoxAddi[2]);
            this.groupBox3.Controls.Add(this.textBoxAddi[1]);
            this.groupBox3.Controls.Add(this.textBoxAddi[0]);
            this.groupBox3.Controls.Add(this.labelAddi[3]);
            this.groupBox3.Controls.Add(this.labelAddi[2]);
            this.groupBox3.Controls.Add(this.labelAddi[1]);
            this.groupBox3.Controls.Add(this.labelAddi[0]);



            iMode = intMode;
            switch (intMode)
            {
                case 1: //增加人员
                    this.btnAdd.Visible = true;
                    //this.btnEdit.Visible = false;
                    this.btnDelPic.Visible = false;
                    break;
                case 2: //修改人员
                    this.btnAdd.Visible = false;
                    //this.btnEdit.Visible = true;
                    this.Text = "修改人员信息";
                    break;
                default:
                    break;

            }
        }



        private void textBoxUserName_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxUsername.Text.Length > 10)
            {
                this.errorProviderPeople.SetError(this.textBoxUsername, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void formPeopleNow_Load(object sender, EventArgs e)
        {
            
            int i;
            this.CancelButton = this.btnCancel;
            //版本处理
            switch (strVersion)
            {
                case "1": //基础版
                    btnFind.Enabled = false;
                    break;
                case "2": //标准版
                    break;
                case "3": //豪华版
                    break;
                default:
                    break;
            } 

            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;
            //附加人员信息
            sqlConn.Open();
            sqlComm.CommandText = "SELECT 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 系统参数表";
            if (dSet.Tables.Contains("系统参数表")) dSet.Tables.Remove("系统参数表");
            sqlDA.Fill(dSet, "系统参数表");
            sqlConn.Close();

            if (dSet.Tables["系统参数表"].Rows.Count < 1)
            {
                for (i = 0; i < 4; i++)
                {
                    labelAddi[i].Text = "附加属性:";
                    textBoxAddi[i].Text = "";
                    textBoxAddi[i].Enabled = false;
                }
            }
            else
            {
                for (i = 0; i < 4; i++)
                {
                    if (dSet.Tables["系统参数表"].Rows[0][i].ToString() == "")
                    {
                        labelAddi[i].Text = "附加属性:";
                        textBoxAddi[i].Text = "";
                        textBoxAddi[i].Enabled = false;
                    }
                    else
                    {
                        labelAddi[i].Text = dSet.Tables["系统参数表"].Rows[0][i].ToString()+":";
                        textBoxAddi[i].Text = "";
                        textBoxAddi[i].Enabled = true;
                    }
                }

            }

            if (iMode == 2) //修改模式
            {
                initFrmPeople();
                //checkBoxSignAll.Visible = false;
            }
            else
            {
                sqlConn.Open();
                sqlComm.CommandText = "SELECT 分组名称 FROM 人员分组表";
                if (dSet.Tables.Contains("人员分组表")) dSet.Tables.Remove("人员分组表");
                sqlDA.Fill(dSet, "人员分组表");
                sqlConn.Close();

                for (i = 0; i < dSet.Tables["人员分组表"].Rows.Count; i++)
                {
                    comboBoxGroup.Items.Add(dSet.Tables["人员分组表"].Rows[i][0]);
                }
            }

        }

        //初始化窗口
        private void initFrmPeople()
        {
            int i;
            //
            System.Data.SqlClient.SqlDataReader sqldr;

            sqlConn.Open();


            sqlComm.CommandText = "SELECT 姓名, 性别, 年龄, 职称, 职务, 所属单位, 科室, 手机, 电话, 传真, EMail, 通讯地址, 邮政编码, 主卡号, 人员类别, 附加属性一, 附加属性二, 附加属性三, 附加属性四 FROM 人员表 WHERE (人员ID = " + intPID.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            if (!sqldr.HasRows)
            {
                MessageBox.Show("数据库错误！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;

            }

            sqldr.Read();
            textBoxUsername.Text = sqldr.GetValue(0).ToString();
            comboBoxGender.Text = sqldr.GetValue(1).ToString();
            textBoxAge.Text = sqldr.GetValue(2).ToString();
            textBoxPPost.Text = sqldr.GetValue(3).ToString();
            textBoxPost.Text = sqldr.GetValue(4).ToString();
            textBoxCompany.Text = sqldr.GetValue(5).ToString();
            textBoxUnit.Text = sqldr.GetValue(6).ToString();
            textBoxMobi.Text = sqldr.GetValue(7).ToString();
            textBoxTelephone.Text = sqldr.GetValue(8).ToString();
            textBoxFax.Text = sqldr.GetValue(9).ToString();
            textBoxEmail.Text = sqldr.GetValue(10).ToString();
            textBoxAddress.Text = sqldr.GetValue(11).ToString();
            textBoxPostcode.Text = sqldr.GetValue(12).ToString();
            textBoxCard.Text = sqldr.GetValue(13).ToString();
            comboBoxGroup.Text = sqldr.GetValue(14).ToString();
            for (i = 0; i < 4; i++)
            {
                if(textBoxAddi[i].Enabled)
                    textBoxAddi[i].Text = sqldr.GetValue(15+i).ToString();
            }
            sqldr.Close();

            sqlComm.CommandText = "SELECT 照片 FROM 照片表 WHERE (人员ID = "+intPID.ToString()+")";
            sqlDA.SelectCommand = sqlComm;
            if (dSet.Tables.Contains("PHOTO")) dSet.Tables.Remove("PHOTO");
            sqlDA.Fill(dSet, "PHOTO");

            sqlComm.CommandText = "SELECT 分组名称 FROM 人员分组表";
            if (dSet.Tables.Contains("人员分组表")) dSet.Tables.Remove("人员分组表");
            sqlDA.Fill(dSet, "人员分组表");

            try
            {
                if (dSet.Tables["PHOTO"].Rows.Count != 0)
                {
                    byte[] bytePhoto = (byte[])dSet.Tables["PHOTO"].Rows[0][0];
                    MemoryStream StreamPhoto = new MemoryStream(bytePhoto);
                    this.pictureBoxPhoto.Image = Image.FromStream(StreamPhoto);
                }

                for (i = 0; i < dSet.Tables["人员分组表"].Rows.Count;i++ )
                {
                    comboBoxGroup.Items.Add(dSet.Tables["人员分组表"].Rows[i][0]);
                }
                
            }
            catch
            {
                sqlConn.Close();
                return;
            }


            sqlConn.Close();


        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            if (openFileDialogPhoto.ShowDialog() == DialogResult.OK)
            {
                textBoxPhoto.Text = openFileDialogPhoto.FileName;
                //
                pictureBoxPhoto.Image = System.Drawing.Image.FromFile(textBoxPhoto.Text);
            }
            
        }

        //人员加入
        private void btnAdd_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlDataReader sqldr;
            System.Data.SqlClient.SqlTransaction sqlta;
            int i;
            
            
            //姓名不能为空
            if (textBoxUsername.Text == "")
            {
                MessageBox.Show("姓名不正确！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //卡号不能为空
            if (textBoxCard.Text == "")
            {
                MessageBox.Show("尚未刷卡！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.Connection = sqlConn;

            if (textBoxCard.Text.Trim() != "")
            {

                sqlComm.CommandText = "SELECT 人员ID, 姓名 FROM 人员表 WHERE (主卡号 = N'" + textBoxCard.Text.Trim() + "')";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    MessageBox.Show("智能卡号重复，使用人为：" + sqldr.GetValue(1).ToString(), "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sqldr.Close();
                    sqlConn.Close();
                    return;
                }
                sqldr.Close();
            }


            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            string strAge;
            strAge = textBoxAge.Text.Trim();
            if (strAge == "") strAge = "NULL";


            try
            {
                sqlComm.CommandText = "INSERT INTO 人员表 (姓名, 性别, 年龄, 职称, 职务, 所属单位, 科室, 手机, 电话, 传真, EMail, 通讯地址, 邮政编码, 主卡号, 人员类别 , 附加属性一, 附加属性二, 附加属性三, 附加属性四) VALUES (N'" + textBoxUsername.Text.Trim() + "', N'" + comboBoxGender.Text + "', " + strAge + ", N'" + textBoxPPost.Text.Trim() + "', N'" + textBoxPost.Text.Trim() + "', N'" + textBoxCompany.Text.Trim() + "', N'" + textBoxUnit.Text.Trim() + "', '" + textBoxMobi.Text.Trim() + "', '" + textBoxTelephone.Text.Trim() + "', '" + textBoxFax.Text.Trim() + "', '" + textBoxEmail.Text.Trim() + "',N'" + textBoxAddress.Text.Trim() + "', '" + textBoxPostcode.Text.Trim() + "', N'" + textBoxCard.Text.Trim() + "',N'" + comboBoxGroup.Text + "', N'" + textBoxAddi[0].Text.Trim() + "', N'" + textBoxAddi[1].Text.Trim() + "', N'" + textBoxAddi[2].Text.Trim() + "', N'" + textBoxAddi[3].Text.Trim() + "')";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "SELECT @@IDENTITY";
                sqldr = sqlComm.ExecuteReader();
                sqldr.Read();
                string strPId = sqldr.GetValue(0).ToString();
                sqldr.Close();

                if (textBoxPhoto.Text.Trim() != "")
                {

                    //照片表
                    FileStream photoStream = new FileStream(textBoxPhoto.Text, FileMode.Open, FileAccess.Read);
                    byte[] photoBytes = new byte[photoStream.Length];
                    photoStream.Read(photoBytes, 0, (int)photoStream.Length);
                    photoStream.Close();

                    //提交
                    sqlComm.CommandText = "INSERT INTO 照片表 (人员ID, 照片) VALUES (@PID , @Photo)";
                    sqlComm.Parameters.Clear();
                    sqlComm.Parameters.Add("@PID", SqlDbType.Int);
                    sqlComm.Parameters.Add("@Photo", SqlDbType.Image);
                    sqlComm.Parameters["@PID"].Value = Int32.Parse(strPId);
                    sqlComm.Parameters["@Photo"].Value = photoBytes;
   
                    sqlComm.ExecuteNonQuery();

                }

                //参加所有会议

                    sqlComm.CommandText = "SELECT 会议ID FROM 会议表";
                    if (dSet.Tables.Contains("会议表")) dSet.Tables.Remove("会议表");
                    sqlDA.Fill(dSet, "会议表");

                    for (i = 0; i < dSet.Tables["会议表"].Rows.Count; i++)
                    {
                        sqlComm.CommandText = "INSERT INTO 参会人员表 (会议ID, 人员ID, 分组名称, 座位号) VALUES (" + dSet.Tables["会议表"].Rows[i][0] + ", " + strPId + ", N'"+textBoxfzmc.Text.Trim()+"', N'"+textBoxzwh.Text.Trim()+"')";
                        sqlComm.ExecuteNonQuery();
                    }

                

                sqlta.Commit();
            }
            catch (Exception err)
            {
                //回滚
                MessageBox.Show("数据库错误！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("人员已加入到数据库！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            sqlConn.Close();


        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlTransaction sqlta;
            System.Data.SqlClient.SqlDataReader sqldr;


            //姓名不能为空
            if (textBoxUsername.Text == "")
            {
                MessageBox.Show("姓名不正确！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.Connection = sqlConn;

            sqlComm.CommandText = "SELECT 人员ID, 姓名 FROM 人员表 WHERE (主卡号 = N'" + textBoxCard.Text.Trim() + "') AND (人员ID <> "+intPID.ToString()+")";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Read();
                MessageBox.Show("智能卡号重复，使用人为：" + sqldr.GetValue(1).ToString(), "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sqldr.Close();
                sqlConn.Close();
                return;
            }
            sqldr.Close();


            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;

            string strAge;
            strAge = textBoxAge.Text.Trim();
            if (strAge == "") strAge = "NULL";


            try
            {
                sqlComm.CommandText = "UPDATE 人员表 SET 姓名 = N'" + textBoxUsername.Text.Trim() + "', 性别 = N'" + comboBoxGender.Text + "', 年龄 = " + strAge + ", 职称 = N'" + textBoxPPost.Text.Trim() + "', 职务 = N'" + textBoxPost.Text.Trim() + "', 所属单位 = N'" + textBoxCompany.Text.Trim() + "', 科室 = N'" + textBoxUnit.Text.Trim() + "', 手机 = '" + textBoxMobi.Text.Trim() + "', 电话 = '" + textBoxTelephone.Text.Trim() + "', 传真 = '" + textBoxFax.Text.Trim() + "', EMail = '" + textBoxEmail.Text.Trim() + "', 通讯地址 = N'" + textBoxAddress.Text.Trim() + "', 邮政编码 = '" + textBoxPostcode.Text.Trim() + "', 主卡号 = N'" + textBoxCard.Text.Trim() + "', 人员类别 = N'" + comboBoxGroup.Text.Trim() + "', 附加属性一 = N'" + textBoxAddi[0].Text.Trim() + "', 附加属性二 = N'" + textBoxAddi[1].Text.Trim() + "', 附加属性三 = N'" + textBoxAddi[2].Text.Trim() + "', 附加属性四 = N'" + textBoxAddi[3].Text.Trim() + "' WHERE (人员ID = " + intPID.ToString() + ")";
                sqlComm.ExecuteNonQuery();


                if(boolDelPic)
                {
                    sqlComm.CommandText = "DELETE FROM 照片表 WHERE (人员ID = "+intPID.ToString()+")";
                    sqlComm.ExecuteNonQuery();
                }

                if (textBoxPhoto.Text.Trim() != "")
                {
                    if (!boolDelPic)
                    {
                        sqlComm.CommandText = "DELETE FROM 照片表 WHERE (人员ID = " + intPID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();
                    }

                    //照片表
                    FileStream photoStream = new FileStream(textBoxPhoto.Text, FileMode.Open, FileAccess.Read);
                    byte[] photoBytes = new byte[photoStream.Length];
                    photoStream.Read(photoBytes, 0, (int)photoStream.Length);
                    photoStream.Close();

                    //提交
                    sqlComm.CommandText = "INSERT INTO 照片表 (人员ID, 照片) VALUES (@PID , @Photo)";
                    if (!sqlComm.Parameters.Contains("@PID"))
                        sqlComm.Parameters.Add("@PID", SqlDbType.Int);
                    if (!sqlComm.Parameters.Contains("@Photo")) 
                        sqlComm.Parameters.Add("@Photo", SqlDbType.Image);
                    sqlComm.Parameters["@PID"].Value = intPID;
                    sqlComm.Parameters["@Photo"].Value = photoBytes;

                    sqlComm.ExecuteNonQuery();

                }
                sqlta.Commit();
            }

            catch (Exception err)
            {
                //回滚
                MessageBox.Show("数据库错误！", "数据库错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();

                return;

            }
            finally
            {
                sqlConn.Close();
            }
            MessageBox.Show("人员已修改！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

           
        }

        private void btnDelPic_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否要删除人员照片？该过程不可回退!", "信息提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                return;

            boolDelPic = true;
            this.pictureBoxPhoto.Image = null;

        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            this.textBoxCard.Focus();
            this.textBoxCard.SelectAll();
        }

        private void textBoxCard_TextChanged(object sender, EventArgs e)
        {
           //this.textBoxCard.SelectAll();
            if(this.textBoxCard.Text.Length==intCardLength)
                this.textBoxCard.SelectAll();
        }

        private void textBoxAge_Validating(object sender, CancelEventArgs e)
        {
            int intOut = 0;
            if (Int32.TryParse(textBoxAge.Text, out intOut) || textBoxAge.Text == "")
            {
                this.errorProviderPeople.Clear();
            }
            else
            {
                this.errorProviderPeople.SetError(this.textBoxAge, "请输入正确的年龄");
                e.Cancel = true;
            }
        }

        private void textBoxPPost_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxPPost.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxPPost, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxPost_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxPost.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxPost, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxCompany_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxCompany.Text.Length > 100)
            {
                this.errorProviderPeople.SetError(this.textBoxCompany, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxUnit_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxUnit.Text.Length > 50)
            {
                this.errorProviderPeople.SetError(this.textBoxUnit, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxAddress_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxAddress.Text.Length > 50)
            {
                this.errorProviderPeople.SetError(this.textBoxAddress, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxPostcode_Validating(object sender, CancelEventArgs e)
        {
            System.Text.RegularExpressions.Regex rExpression = new System.Text.RegularExpressions.Regex(@"^\d{6}$");

            if (rExpression.IsMatch(textBoxPostcode.Text) || textBoxPostcode.Text == "")
            {
                this.errorProviderPeople.Clear();
            }
            else
            {
                this.errorProviderPeople.SetError(this.textBoxPostcode, "输入正确的邮政编码");
                e.Cancel = true;
            }
        }

        private void textBoxTelephone_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxTelephone.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxTelephone, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxMobi_Validating(object sender, CancelEventArgs e)
        {
            System.Text.RegularExpressions.Regex rExpression = new System.Text.RegularExpressions.Regex(@"^\d{11}$");

            if (rExpression.IsMatch(textBoxMobi.Text) || textBoxMobi.Text == "")
            {
                this.errorProviderPeople.Clear();
            }
            else
            {
                this.errorProviderPeople.SetError(this.textBoxMobi, "输入正确的手机号码");
                e.Cancel = true;
            }
        }

        private void textBoxFax_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxFax.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxFax, "字段过长");
                e.Cancel = true;
            }
            else
            {
                this.errorProviderPeople.Clear();
            }
        }

        private void textBoxEmail_Validating(object sender, CancelEventArgs e)
        {
            System.Text.RegularExpressions.Regex rExpression = new System.Text.RegularExpressions.Regex(@"^([\w\-]+\.)*[\w\-]+@([\w\-]+\.)+([\w\-]{2,3})$");

            if (rExpression.IsMatch(textBoxEmail.Text) || textBoxEmail.Text == "")
            {
                this.errorProviderPeople.Clear();
            }
            else
            {
                this.errorProviderPeople.SetError(this.textBoxEmail, "输入正确的电子邮件");
                e.Cancel = true;
            }
        }
    }
}