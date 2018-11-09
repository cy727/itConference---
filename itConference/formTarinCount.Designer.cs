namespace itConference
{
    partial class formTarinCount
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formTarinCount));
            this.btnOutput = new System.Windows.Forms.Button();
            this.listBoxInfo = new System.Windows.Forms.ListBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataGridViewUnSign = new System.Windows.Forms.DataGridView();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnCount = new System.Windows.Forms.Button();
            this.dataGridViewNew = new System.Windows.Forms.DataGridView();
            this.dataGridViewSign = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.openFileDialogoff = new System.Windows.Forms.OpenFileDialog();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.saveFileDialogOutput = new System.Windows.Forms.SaveFileDialog();
            this.sqlDA = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelConference = new System.Windows.Forms.Label();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.listBoxInput = new System.Windows.Forms.ListBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.toolStripStatusLabelCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusStripCount = new System.Windows.Forms.StatusStrip();
            this.btnOffInput = new System.Windows.Forms.Button();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewUnSign)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNew)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSign)).BeginInit();
            this.tabPage3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.statusStripCount.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOutput
            // 
            this.btnOutput.Location = new System.Drawing.Point(260, 552);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(75, 23);
            this.btnOutput.TabIndex = 13;
            this.btnOutput.Text = "数据导出";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.btnOutput_Click);
            // 
            // listBoxInfo
            // 
            this.listBoxInfo.FormattingEnabled = true;
            this.listBoxInfo.Items.AddRange(new object[] {
            "性别",
            "年龄",
            "职称",
            "职务",
            "所属单位",
            "科室",
            "手机",
            "电话",
            "传真",
            "EMail",
            "通讯地址",
            "邮政编码",
            "人员类别",
            "附加属性一",
            "附加属性二",
            "附加属性三",
            "附加属性四",
            "编号"});
            this.listBoxInfo.Location = new System.Drawing.Point(7, 20);
            this.listBoxInfo.Name = "listBoxInfo";
            this.listBoxInfo.Size = new System.Drawing.Size(144, 238);
            this.listBoxInfo.TabIndex = 0;
            this.listBoxInfo.DoubleClick += new System.EventHandler(this.listBoxInfo_DoubleClick);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dataGridViewUnSign);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(531, 167);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "缺席";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataGridViewUnSign
            // 
            this.dataGridViewUnSign.AllowUserToAddRows = false;
            this.dataGridViewUnSign.AllowUserToDeleteRows = false;
            this.dataGridViewUnSign.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewUnSign.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewUnSign.Location = new System.Drawing.Point(3, 3);
            this.dataGridViewUnSign.Name = "dataGridViewUnSign";
            this.dataGridViewUnSign.ReadOnly = true;
            this.dataGridViewUnSign.Size = new System.Drawing.Size(525, 161);
            this.dataGridViewUnSign.TabIndex = 2;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(344, 552);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnCount
            // 
            this.btnCount.Location = new System.Drawing.Point(23, 552);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(75, 23);
            this.btnCount.TabIndex = 12;
            this.btnCount.Text = "数据统计";
            this.btnCount.UseVisualStyleBackColor = true;
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // dataGridViewNew
            // 
            this.dataGridViewNew.AllowUserToAddRows = false;
            this.dataGridViewNew.AllowUserToDeleteRows = false;
            this.dataGridViewNew.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewNew.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewNew.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewNew.Name = "dataGridViewNew";
            this.dataGridViewNew.ReadOnly = true;
            this.dataGridViewNew.Size = new System.Drawing.Size(531, 167);
            this.dataGridViewNew.TabIndex = 2;
            // 
            // dataGridViewSign
            // 
            this.dataGridViewSign.AllowUserToAddRows = false;
            this.dataGridViewSign.AllowUserToDeleteRows = false;
            this.dataGridViewSign.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSign.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewSign.Location = new System.Drawing.Point(3, 3);
            this.dataGridViewSign.Name = "dataGridViewSign";
            this.dataGridViewSign.ReadOnly = true;
            this.dataGridViewSign.Size = new System.Drawing.Size(386, 161);
            this.dataGridViewSign.TabIndex = 1;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dataGridViewNew);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(531, 167);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "新增";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // openFileDialogoff
            // 
            this.openFileDialogoff.RestoreDirectory = true;
            // 
            // sqlConn
            // 
            this.sqlConn.FireInfoMessageEventOnUserErrors = false;
            // 
            // saveFileDialogOutput
            // 
            this.saveFileDialogOutput.CreatePrompt = true;
            this.saveFileDialogOutput.DefaultExt = "xls";
            this.saveFileDialogOutput.Filter = "excel文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            this.saveFileDialogOutput.RestoreDirectory = true;
            // 
            // sqlDA
            // 
            this.sqlDA.DeleteCommand = this.sqlDeleteCommand1;
            this.sqlDA.InsertCommand = this.sqlInsertCommand1;
            this.sqlDA.SelectCommand = this.sqlSelectCommand1;
            this.sqlDA.UpdateCommand = this.sqlUpdateCommand1;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Azure;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.labelConference);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(432, 53);
            this.panel1.TabIndex = 15;
            // 
            // labelConference
            // 
            this.labelConference.AutoSize = true;
            this.labelConference.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelConference.Location = new System.Drawing.Point(11, 20);
            this.labelConference.Name = "labelConference";
            this.labelConference.Size = new System.Drawing.Size(0, 19);
            this.labelConference.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridViewSign);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(392, 167);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "签到";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(3, 16);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(400, 193);
            this.tabControl1.TabIndex = 0;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.tabControl1);
            this.groupBox3.Location = new System.Drawing.Point(13, 337);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(406, 212);
            this.groupBox3.TabIndex = 11;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "签到明细";
            // 
            // btnDown
            // 
            this.btnDown.Location = new System.Drawing.Point(162, 153);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(75, 23);
            this.btnDown.TabIndex = 5;
            this.btnDown.Text = "下移";
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Location = new System.Drawing.Point(162, 124);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(75, 23);
            this.btnUp.TabIndex = 4;
            this.btnUp.Text = "上移";
            this.btnUp.UseVisualStyleBackColor = true;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(161, 70);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "<<";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // listBoxInput
            // 
            this.listBoxInput.FormattingEnabled = true;
            this.listBoxInput.Items.AddRange(new object[] {
            "姓名"});
            this.listBoxInput.Location = new System.Drawing.Point(247, 20);
            this.listBoxInput.Name = "listBoxInput";
            this.listBoxInput.Size = new System.Drawing.Size(144, 238);
            this.listBoxInput.TabIndex = 1;
            this.listBoxInput.DoubleClick += new System.EventHandler(this.listBoxInput_DoubleClick);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(161, 40);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = ">>";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnDown);
            this.groupBox1.Controls.Add(this.btnUp);
            this.groupBox1.Controls.Add(this.btnDel);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.listBoxInput);
            this.groupBox1.Controls.Add(this.listBoxInfo);
            this.groupBox1.Location = new System.Drawing.Point(13, 57);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(406, 274);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "人员信息";
            // 
            // toolStripStatusLabelCount
            // 
            this.toolStripStatusLabelCount.Name = "toolStripStatusLabelCount";
            this.toolStripStatusLabelCount.Size = new System.Drawing.Size(0, 17);
            // 
            // statusStripCount
            // 
            this.statusStripCount.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelCount});
            this.statusStripCount.Location = new System.Drawing.Point(0, 590);
            this.statusStripCount.Name = "statusStripCount";
            this.statusStripCount.Size = new System.Drawing.Size(432, 22);
            this.statusStripCount.TabIndex = 10;
            this.statusStripCount.Text = "statusStrip1";
            // 
            // btnOffInput
            // 
            this.btnOffInput.Location = new System.Drawing.Point(104, 552);
            this.btnOffInput.Name = "btnOffInput";
            this.btnOffInput.Size = new System.Drawing.Size(117, 23);
            this.btnOffInput.TabIndex = 2;
            this.btnOffInput.Text = "离线签到数据导入";
            this.btnOffInput.UseVisualStyleBackColor = true;
            this.btnOffInput.Click += new System.EventHandler(this.btnOffInput_Click);
            // 
            // formTarinCount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(432, 612);
            this.Controls.Add(this.btnOutput);
            this.Controls.Add(this.btnOffInput);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStripCount);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "formTarinCount";
            this.Text = "会议统计";
            this.Load += new System.EventHandler(this.formTarinCount_Load);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewUnSign)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNew)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSign)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.statusStripCount.ResumeLayout(false);
            this.statusStripCount.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOutput;
        private System.Windows.Forms.ListBox listBoxInfo;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dataGridViewUnSign;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnCount;
        private System.Windows.Forms.DataGridView dataGridViewNew;
        private System.Windows.Forms.DataGridView dataGridViewSign;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Windows.Forms.OpenFileDialog openFileDialogoff;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Windows.Forms.SaveFileDialog saveFileDialogOutput;
        private System.Data.SqlClient.SqlDataAdapter sqlDA;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelConference;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnDown;
        private System.Windows.Forms.Button btnUp;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.ListBox listBoxInput;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelCount;
        private System.Windows.Forms.StatusStrip statusStripCount;
        private System.Windows.Forms.Button btnOffInput;
    }
}