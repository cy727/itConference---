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
            this.labelAddi[0].Text = "��������һ��";
            // 
            // labelAddi1
            // 
            this.labelAddi[1].AutoSize = true;
            this.labelAddi[1].Location = new System.Drawing.Point(244, 21);
            this.labelAddi[1].Name = "labelAddi1";
            this.labelAddi[1].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[1].TabIndex = 1;
            this.labelAddi[1].Text = "�������Զ���";
            // 
            // labelAddi2
            // 
            this.labelAddi[2].AutoSize = true;
            this.labelAddi[2].Location = new System.Drawing.Point(8, 48);
            this.labelAddi[2].Name = "labelAddi2";
            this.labelAddi[2].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[2].TabIndex = 2;
            this.labelAddi[2].Text = "������������";
            // 
            // labelAddi3
            // 
            this.labelAddi[3].AutoSize = true;
            this.labelAddi[3].Location = new System.Drawing.Point(244, 48);
            this.labelAddi[3].Name = "labelAddi3";
            this.labelAddi[3].Size = new System.Drawing.Size(77, 12);
            this.labelAddi[3].TabIndex = 3;
            this.labelAddi[3].Text = "���������ģ�";
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
                case 1: //������Ա
                    this.btnAdd.Visible = true;
                    //this.btnEdit.Visible = false;
                    this.btnDelPic.Visible = false;
                    break;
                case 2: //�޸���Ա
                    this.btnAdd.Visible = false;
                    //this.btnEdit.Visible = true;
                    this.Text = "�޸���Ա��Ϣ";
                    break;
                default:
                    break;

            }
        }



        private void textBoxUserName_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxUsername.Text.Length > 10)
            {
                this.errorProviderPeople.SetError(this.textBoxUsername, "�ֶι���");
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
            //�汾����
            switch (strVersion)
            {
                case "1": //������
                    btnFind.Enabled = false;
                    break;
                case "2": //��׼��
                    break;
                case "3": //������
                    break;
                default:
                    break;
            } 

            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;
            //������Ա��Ϣ
            sqlConn.Open();
            sqlComm.CommandText = "SELECT ��������һ, �������Զ�, ����������, ���������� FROM ϵͳ������";
            if (dSet.Tables.Contains("ϵͳ������")) dSet.Tables.Remove("ϵͳ������");
            sqlDA.Fill(dSet, "ϵͳ������");
            sqlConn.Close();

            if (dSet.Tables["ϵͳ������"].Rows.Count < 1)
            {
                for (i = 0; i < 4; i++)
                {
                    labelAddi[i].Text = "��������:";
                    textBoxAddi[i].Text = "";
                    textBoxAddi[i].Enabled = false;
                }
            }
            else
            {
                for (i = 0; i < 4; i++)
                {
                    if (dSet.Tables["ϵͳ������"].Rows[0][i].ToString() == "")
                    {
                        labelAddi[i].Text = "��������:";
                        textBoxAddi[i].Text = "";
                        textBoxAddi[i].Enabled = false;
                    }
                    else
                    {
                        labelAddi[i].Text = dSet.Tables["ϵͳ������"].Rows[0][i].ToString()+":";
                        textBoxAddi[i].Text = "";
                        textBoxAddi[i].Enabled = true;
                    }
                }

            }

            if (iMode == 2) //�޸�ģʽ
            {
                initFrmPeople();
                //checkBoxSignAll.Visible = false;
            }
            else
            {
                sqlConn.Open();
                sqlComm.CommandText = "SELECT �������� FROM ��Ա�����";
                if (dSet.Tables.Contains("��Ա�����")) dSet.Tables.Remove("��Ա�����");
                sqlDA.Fill(dSet, "��Ա�����");
                sqlConn.Close();

                for (i = 0; i < dSet.Tables["��Ա�����"].Rows.Count; i++)
                {
                    comboBoxGroup.Items.Add(dSet.Tables["��Ա�����"].Rows[i][0]);
                }
            }

        }

        //��ʼ������
        private void initFrmPeople()
        {
            int i;
            //
            System.Data.SqlClient.SqlDataReader sqldr;

            sqlConn.Open();


            sqlComm.CommandText = "SELECT ����, �Ա�, ����, ְ��, ְ��, ������λ, ����, �ֻ�, �绰, ����, EMail, ͨѶ��ַ, ��������, ������, ��Ա���, ��������һ, �������Զ�, ����������, ���������� FROM ��Ա�� WHERE (��ԱID = " + intPID.ToString() + ")";
            sqldr = sqlComm.ExecuteReader();
            if (!sqldr.HasRows)
            {
                MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            sqlComm.CommandText = "SELECT ��Ƭ FROM ��Ƭ�� WHERE (��ԱID = "+intPID.ToString()+")";
            sqlDA.SelectCommand = sqlComm;
            if (dSet.Tables.Contains("PHOTO")) dSet.Tables.Remove("PHOTO");
            sqlDA.Fill(dSet, "PHOTO");

            sqlComm.CommandText = "SELECT �������� FROM ��Ա�����";
            if (dSet.Tables.Contains("��Ա�����")) dSet.Tables.Remove("��Ա�����");
            sqlDA.Fill(dSet, "��Ա�����");

            try
            {
                if (dSet.Tables["PHOTO"].Rows.Count != 0)
                {
                    byte[] bytePhoto = (byte[])dSet.Tables["PHOTO"].Rows[0][0];
                    MemoryStream StreamPhoto = new MemoryStream(bytePhoto);
                    this.pictureBoxPhoto.Image = Image.FromStream(StreamPhoto);
                }

                for (i = 0; i < dSet.Tables["��Ա�����"].Rows.Count;i++ )
                {
                    comboBoxGroup.Items.Add(dSet.Tables["��Ա�����"].Rows[i][0]);
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

        //��Ա����
        private void btnAdd_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlDataReader sqldr;
            System.Data.SqlClient.SqlTransaction sqlta;
            int i;
            
            
            //��������Ϊ��
            if (textBoxUsername.Text == "")
            {
                MessageBox.Show("��������ȷ��", "�������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //���Ų���Ϊ��
            if (textBoxCard.Text == "")
            {
                MessageBox.Show("��δˢ����", "�������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.Connection = sqlConn;

            if (textBoxCard.Text.Trim() != "")
            {

                sqlComm.CommandText = "SELECT ��ԱID, ���� FROM ��Ա�� WHERE (������ = N'" + textBoxCard.Text.Trim() + "')";
                sqldr = sqlComm.ExecuteReader();

                if (sqldr.HasRows)
                {
                    sqldr.Read();
                    MessageBox.Show("���ܿ����ظ���ʹ����Ϊ��" + sqldr.GetValue(1).ToString(), "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                sqlComm.CommandText = "INSERT INTO ��Ա�� (����, �Ա�, ����, ְ��, ְ��, ������λ, ����, �ֻ�, �绰, ����, EMail, ͨѶ��ַ, ��������, ������, ��Ա��� , ��������һ, �������Զ�, ����������, ����������) VALUES (N'" + textBoxUsername.Text.Trim() + "', N'" + comboBoxGender.Text + "', " + strAge + ", N'" + textBoxPPost.Text.Trim() + "', N'" + textBoxPost.Text.Trim() + "', N'" + textBoxCompany.Text.Trim() + "', N'" + textBoxUnit.Text.Trim() + "', '" + textBoxMobi.Text.Trim() + "', '" + textBoxTelephone.Text.Trim() + "', '" + textBoxFax.Text.Trim() + "', '" + textBoxEmail.Text.Trim() + "',N'" + textBoxAddress.Text.Trim() + "', '" + textBoxPostcode.Text.Trim() + "', N'" + textBoxCard.Text.Trim() + "',N'" + comboBoxGroup.Text + "', N'" + textBoxAddi[0].Text.Trim() + "', N'" + textBoxAddi[1].Text.Trim() + "', N'" + textBoxAddi[2].Text.Trim() + "', N'" + textBoxAddi[3].Text.Trim() + "')";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "SELECT @@IDENTITY";
                sqldr = sqlComm.ExecuteReader();
                sqldr.Read();
                string strPId = sqldr.GetValue(0).ToString();
                sqldr.Close();

                if (textBoxPhoto.Text.Trim() != "")
                {

                    //��Ƭ��
                    FileStream photoStream = new FileStream(textBoxPhoto.Text, FileMode.Open, FileAccess.Read);
                    byte[] photoBytes = new byte[photoStream.Length];
                    photoStream.Read(photoBytes, 0, (int)photoStream.Length);
                    photoStream.Close();

                    //�ύ
                    sqlComm.CommandText = "INSERT INTO ��Ƭ�� (��ԱID, ��Ƭ) VALUES (@PID , @Photo)";
                    sqlComm.Parameters.Clear();
                    sqlComm.Parameters.Add("@PID", SqlDbType.Int);
                    sqlComm.Parameters.Add("@Photo", SqlDbType.Image);
                    sqlComm.Parameters["@PID"].Value = Int32.Parse(strPId);
                    sqlComm.Parameters["@Photo"].Value = photoBytes;
   
                    sqlComm.ExecuteNonQuery();

                }

                //�μ����л���

                    sqlComm.CommandText = "SELECT ����ID FROM �����";
                    if (dSet.Tables.Contains("�����")) dSet.Tables.Remove("�����");
                    sqlDA.Fill(dSet, "�����");

                    for (i = 0; i < dSet.Tables["�����"].Rows.Count; i++)
                    {
                        sqlComm.CommandText = "INSERT INTO �λ���Ա�� (����ID, ��ԱID, ��������, ��λ��) VALUES (" + dSet.Tables["�����"].Rows[i][0] + ", " + strPId + ", N'"+textBoxfzmc.Text.Trim()+"', N'"+textBoxzwh.Text.Trim()+"')";
                        sqlComm.ExecuteNonQuery();
                    }

                

                sqlta.Commit();
            }
            catch (Exception err)
            {
                //�ع�
                MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                sqlConn.Close();
                return;

            }
            finally
            {
                MessageBox.Show("��Ա�Ѽ��뵽���ݿ⣡", "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            sqlConn.Close();


        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlTransaction sqlta;
            System.Data.SqlClient.SqlDataReader sqldr;


            //��������Ϊ��
            if (textBoxUsername.Text == "")
            {
                MessageBox.Show("��������ȷ��", "�������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sqlConn.Open();
            sqlComm.Connection = sqlConn;

            sqlComm.CommandText = "SELECT ��ԱID, ���� FROM ��Ա�� WHERE (������ = N'" + textBoxCard.Text.Trim() + "') AND (��ԱID <> "+intPID.ToString()+")";
            sqldr = sqlComm.ExecuteReader();

            if (sqldr.HasRows)
            {
                sqldr.Read();
                MessageBox.Show("���ܿ����ظ���ʹ����Ϊ��" + sqldr.GetValue(1).ToString(), "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                sqlComm.CommandText = "UPDATE ��Ա�� SET ���� = N'" + textBoxUsername.Text.Trim() + "', �Ա� = N'" + comboBoxGender.Text + "', ���� = " + strAge + ", ְ�� = N'" + textBoxPPost.Text.Trim() + "', ְ�� = N'" + textBoxPost.Text.Trim() + "', ������λ = N'" + textBoxCompany.Text.Trim() + "', ���� = N'" + textBoxUnit.Text.Trim() + "', �ֻ� = '" + textBoxMobi.Text.Trim() + "', �绰 = '" + textBoxTelephone.Text.Trim() + "', ���� = '" + textBoxFax.Text.Trim() + "', EMail = '" + textBoxEmail.Text.Trim() + "', ͨѶ��ַ = N'" + textBoxAddress.Text.Trim() + "', �������� = '" + textBoxPostcode.Text.Trim() + "', ������ = N'" + textBoxCard.Text.Trim() + "', ��Ա��� = N'" + comboBoxGroup.Text.Trim() + "', ��������һ = N'" + textBoxAddi[0].Text.Trim() + "', �������Զ� = N'" + textBoxAddi[1].Text.Trim() + "', ���������� = N'" + textBoxAddi[2].Text.Trim() + "', ���������� = N'" + textBoxAddi[3].Text.Trim() + "' WHERE (��ԱID = " + intPID.ToString() + ")";
                sqlComm.ExecuteNonQuery();


                if(boolDelPic)
                {
                    sqlComm.CommandText = "DELETE FROM ��Ƭ�� WHERE (��ԱID = "+intPID.ToString()+")";
                    sqlComm.ExecuteNonQuery();
                }

                if (textBoxPhoto.Text.Trim() != "")
                {
                    if (!boolDelPic)
                    {
                        sqlComm.CommandText = "DELETE FROM ��Ƭ�� WHERE (��ԱID = " + intPID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();
                    }

                    //��Ƭ��
                    FileStream photoStream = new FileStream(textBoxPhoto.Text, FileMode.Open, FileAccess.Read);
                    byte[] photoBytes = new byte[photoStream.Length];
                    photoStream.Read(photoBytes, 0, (int)photoStream.Length);
                    photoStream.Close();

                    //�ύ
                    sqlComm.CommandText = "INSERT INTO ��Ƭ�� (��ԱID, ��Ƭ) VALUES (@PID , @Photo)";
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
                //�ع�
                MessageBox.Show("���ݿ����", "���ݿ����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();

                return;

            }
            finally
            {
                sqlConn.Close();
            }
            MessageBox.Show("��Ա���޸ģ�", "��Ϣ��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);

           
        }

        private void btnDelPic_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�Ƿ�Ҫɾ����Ա��Ƭ���ù��̲��ɻ���!", "��Ϣ��ʾ", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
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
                this.errorProviderPeople.SetError(this.textBoxAge, "��������ȷ������");
                e.Cancel = true;
            }
        }

        private void textBoxPPost_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxPPost.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxPPost, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxPost, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxCompany, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxUnit, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxAddress, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxPostcode, "������ȷ����������");
                e.Cancel = true;
            }
        }

        private void textBoxTelephone_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxTelephone.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxTelephone, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxMobi, "������ȷ���ֻ�����");
                e.Cancel = true;
            }
        }

        private void textBoxFax_Validating(object sender, CancelEventArgs e)
        {
            if (this.textBoxFax.Text.Length > 20)
            {
                this.errorProviderPeople.SetError(this.textBoxFax, "�ֶι���");
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
                this.errorProviderPeople.SetError(this.textBoxEmail, "������ȷ�ĵ����ʼ�");
                e.Cancel = true;
            }
        }
    }
}