namespace itConference
{
    partial class formCount
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formCount));
            this.listBoxInfo = new System.Windows.Forms.ListBox();
            this.listBoxInput = new System.Windows.Forms.ListBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxderysl = new System.Windows.Forms.CheckBox();
            this.checkBoxqdrysl = new System.Windows.Forms.CheckBox();
            this.checkBoxqxrysl = new System.Windows.Forms.CheckBox();
            this.checkBoxcdrysl = new System.Windows.Forms.CheckBox();
            this.checkBoxqcsj = new System.Windows.Forms.CheckBox();
            this.checkBoxhztjxx = new System.Windows.Forms.CheckBox();
            this.checkBoxcdrylb = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnSCJ = new System.Windows.Forms.Button();
            this.checkBoxNew = new System.Windows.Forms.CheckBox();
            this.checkBoxqxrylb = new System.Windows.Forms.CheckBox();
            this.btnOffInput = new System.Windows.Forms.Button();
            this.statusStripCount = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataGridViewSign = new System.Windows.Forms.DataGridView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataGridViewUnSign = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dataGridViewNew = new System.Windows.Forms.DataGridView();
            this.btnCount = new System.Windows.Forms.Button();
            this.btnOutput = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDA = new System.Data.SqlClient.SqlDataAdapter();
            this.openFileDialogoff = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialogOutput = new System.Windows.Forms.SaveFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelConference = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.statusStripCount.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSign)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewUnSign)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNew)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listBoxInfo
            // 
            this.listBoxInfo.FormattingEnabled = true;
            this.listBoxInfo.ItemHeight = 12;
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
            this.listBoxInfo.Location = new System.Drawing.Point(7, 18);
            this.listBoxInfo.Name = "listBoxInfo";
            this.listBoxInfo.Size = new System.Drawing.Size(144, 232);
            this.listBoxInfo.TabIndex = 0;
            this.listBoxInfo.DoubleClick += new System.EventHandler(this.listBoxInfo_DoubleClick);
            // 
            // listBoxInput
            // 
            this.listBoxInput.FormattingEnabled = true;
            this.listBoxInput.ItemHeight = 12;
            this.listBoxInput.Items.AddRange(new object[] {
            "姓名"});
            this.listBoxInput.Location = new System.Drawing.Point(247, 18);
            this.listBoxInput.Name = "listBoxInput";
            this.listBoxInput.Size = new System.Drawing.Size(144, 232);
            this.listBoxInput.TabIndex = 1;
            this.listBoxInput.DoubleClick += new System.EventHandler(this.listBoxInput_DoubleClick);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(161, 37);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 21);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = ">>";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(161, 65);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 21);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "<<";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnUp
            // 
            this.btnUp.Location = new System.Drawing.Point(162, 114);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(75, 21);
            this.btnUp.TabIndex = 4;
            this.btnUp.Text = "上移";
            this.btnUp.UseVisualStyleBackColor = true;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // btnDown
            // 
            this.btnDown.Location = new System.Drawing.Point(162, 141);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(75, 21);
            this.btnDown.TabIndex = 5;
            this.btnDown.Text = "下移";
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnDown);
            this.groupBox1.Controls.Add(this.btnUp);
            this.groupBox1.Controls.Add(this.btnDel);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.listBoxInput);
            this.groupBox1.Controls.Add(this.listBoxInfo);
            this.groupBox1.Location = new System.Drawing.Point(13, 53);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(406, 267);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "人员信息";
            // 
            // checkBoxderysl
            // 
            this.checkBoxderysl.AutoSize = true;
            this.checkBoxderysl.Checked = true;
            this.checkBoxderysl.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxderysl.Location = new System.Drawing.Point(17, 18);
            this.checkBoxderysl.Name = "checkBoxderysl";
            this.checkBoxderysl.Size = new System.Drawing.Size(96, 15);
            this.checkBoxderysl.TabIndex = 0;
            this.checkBoxderysl.Text = "登记人员数量";
            this.checkBoxderysl.UseVisualStyleBackColor = true;
            // 
            // checkBoxqdrysl
            // 
            this.checkBoxqdrysl.AutoSize = true;
            this.checkBoxqdrysl.Checked = true;
            this.checkBoxqdrysl.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxqdrysl.Location = new System.Drawing.Point(17, 41);
            this.checkBoxqdrysl.Name = "checkBoxqdrysl";
            this.checkBoxqdrysl.Size = new System.Drawing.Size(96, 15);
            this.checkBoxqdrysl.TabIndex = 1;
            this.checkBoxqdrysl.Text = "签到人员数量";
            this.checkBoxqdrysl.UseVisualStyleBackColor = true;
            // 
            // checkBoxqxrysl
            // 
            this.checkBoxqxrysl.AutoSize = true;
            this.checkBoxqxrysl.Checked = true;
            this.checkBoxqxrysl.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxqxrysl.Location = new System.Drawing.Point(17, 63);
            this.checkBoxqxrysl.Name = "checkBoxqxrysl";
            this.checkBoxqxrysl.Size = new System.Drawing.Size(96, 15);
            this.checkBoxqxrysl.TabIndex = 2;
            this.checkBoxqxrysl.Text = "缺席人员数量";
            this.checkBoxqxrysl.UseVisualStyleBackColor = true;
            // 
            // checkBoxcdrysl
            // 
            this.checkBoxcdrysl.AutoSize = true;
            this.checkBoxcdrysl.Location = new System.Drawing.Point(17, 85);
            this.checkBoxcdrysl.Name = "checkBoxcdrysl";
            this.checkBoxcdrysl.Size = new System.Drawing.Size(96, 15);
            this.checkBoxcdrysl.TabIndex = 3;
            this.checkBoxcdrysl.Text = "迟到人员数量";
            this.checkBoxcdrysl.UseVisualStyleBackColor = true;
            // 
            // checkBoxqcsj
            // 
            this.checkBoxqcsj.AutoSize = true;
            this.checkBoxqcsj.Location = new System.Drawing.Point(17, 107);
            this.checkBoxqcsj.Name = "checkBoxqcsj";
            this.checkBoxqcsj.Size = new System.Drawing.Size(96, 15);
            this.checkBoxqcsj.TabIndex = 4;
            this.checkBoxqcsj.Text = "人员签出时间";
            this.checkBoxqcsj.UseVisualStyleBackColor = true;
            // 
            // checkBoxhztjxx
            // 
            this.checkBoxhztjxx.AutoSize = true;
            this.checkBoxhztjxx.Checked = true;
            this.checkBoxhztjxx.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxhztjxx.Location = new System.Drawing.Point(17, 171);
            this.checkBoxhztjxx.Name = "checkBoxhztjxx";
            this.checkBoxhztjxx.Size = new System.Drawing.Size(96, 15);
            this.checkBoxhztjxx.TabIndex = 6;
            this.checkBoxhztjxx.Text = "汇总统计信息";
            this.checkBoxhztjxx.UseVisualStyleBackColor = true;
            // 
            // checkBoxcdrylb
            // 
            this.checkBoxcdrylb.AutoSize = true;
            this.checkBoxcdrylb.Location = new System.Drawing.Point(17, 128);
            this.checkBoxcdrylb.Name = "checkBoxcdrylb";
            this.checkBoxcdrylb.Size = new System.Drawing.Size(96, 15);
            this.checkBoxcdrylb.TabIndex = 8;
            this.checkBoxcdrylb.Text = "迟到人员列表";
            this.checkBoxcdrylb.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnSCJ);
            this.groupBox2.Controls.Add(this.checkBoxNew);
            this.groupBox2.Controls.Add(this.checkBoxqxrylb);
            this.groupBox2.Controls.Add(this.btnOffInput);
            this.groupBox2.Controls.Add(this.checkBoxcdrylb);
            this.groupBox2.Controls.Add(this.checkBoxhztjxx);
            this.groupBox2.Controls.Add(this.checkBoxqcsj);
            this.groupBox2.Controls.Add(this.checkBoxcdrysl);
            this.groupBox2.Controls.Add(this.checkBoxqxrysl);
            this.groupBox2.Controls.Add(this.checkBoxqdrysl);
            this.groupBox2.Controls.Add(this.checkBoxderysl);
            this.groupBox2.Location = new System.Drawing.Point(425, 53);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(133, 267);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "统计选项";
            // 
            // btnSCJ
            // 
            this.btnSCJ.Location = new System.Drawing.Point(10, 238);
            this.btnSCJ.Name = "btnSCJ";
            this.btnSCJ.Size = new System.Drawing.Size(117, 21);
            this.btnSCJ.TabIndex = 11;
            this.btnSCJ.Text = "手持机数据导入";
            this.btnSCJ.UseVisualStyleBackColor = true;
            this.btnSCJ.Click += new System.EventHandler(this.btnSCJ_Click);
            // 
            // checkBoxNew
            // 
            this.checkBoxNew.AutoSize = true;
            this.checkBoxNew.Location = new System.Drawing.Point(17, 192);
            this.checkBoxNew.Name = "checkBoxNew";
            this.checkBoxNew.Size = new System.Drawing.Size(96, 15);
            this.checkBoxNew.TabIndex = 10;
            this.checkBoxNew.Text = "新增人员统计";
            this.checkBoxNew.UseVisualStyleBackColor = true;
            // 
            // checkBoxqxrylb
            // 
            this.checkBoxqxrylb.AutoSize = true;
            this.checkBoxqxrylb.Checked = true;
            this.checkBoxqxrylb.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxqxrylb.Location = new System.Drawing.Point(17, 150);
            this.checkBoxqxrylb.Name = "checkBoxqxrylb";
            this.checkBoxqxrylb.Size = new System.Drawing.Size(96, 15);
            this.checkBoxqxrylb.TabIndex = 9;
            this.checkBoxqxrylb.Text = "缺席人员列表";
            this.checkBoxqxrylb.UseVisualStyleBackColor = true;
            // 
            // btnOffInput
            // 
            this.btnOffInput.Location = new System.Drawing.Point(10, 211);
            this.btnOffInput.Name = "btnOffInput";
            this.btnOffInput.Size = new System.Drawing.Size(117, 21);
            this.btnOffInput.TabIndex = 2;
            this.btnOffInput.Text = "离线签到数据导入";
            this.btnOffInput.UseVisualStyleBackColor = true;
            this.btnOffInput.Click += new System.EventHandler(this.btnOffInput_Click);
            // 
            // statusStripCount
            // 
            this.statusStripCount.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelCount});
            this.statusStripCount.Location = new System.Drawing.Point(0, 538);
            this.statusStripCount.Name = "statusStripCount";
            this.statusStripCount.Size = new System.Drawing.Size(578, 22);
            this.statusStripCount.TabIndex = 2;
            this.statusStripCount.Text = "statusStrip1";
            // 
            // toolStripStatusLabelCount
            // 
            this.toolStripStatusLabelCount.Name = "toolStripStatusLabelCount";
            this.toolStripStatusLabelCount.Size = new System.Drawing.Size(0, 17);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.tabControl1);
            this.groupBox3.Location = new System.Drawing.Point(13, 325);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(545, 182);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "签到明细";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(3, 15);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(539, 164);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridViewSign);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(531, 138);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "签到";
            this.tabPage1.UseVisualStyleBackColor = true;
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
            this.dataGridViewSign.Size = new System.Drawing.Size(525, 132);
            this.dataGridViewSign.TabIndex = 1;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dataGridViewUnSign);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(531, 138);
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
            this.dataGridViewUnSign.Size = new System.Drawing.Size(525, 149);
            this.dataGridViewUnSign.TabIndex = 2;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dataGridViewNew);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(531, 138);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "新增";
            this.tabPage3.UseVisualStyleBackColor = true;
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
            this.dataGridViewNew.Size = new System.Drawing.Size(531, 154);
            this.dataGridViewNew.TabIndex = 2;
            // 
            // btnCount
            // 
            this.btnCount.Location = new System.Drawing.Point(64, 512);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(75, 21);
            this.btnCount.TabIndex = 4;
            this.btnCount.Text = "数据统计";
            this.btnCount.UseVisualStyleBackColor = true;
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // btnOutput
            // 
            this.btnOutput.Location = new System.Drawing.Point(334, 511);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(75, 21);
            this.btnOutput.TabIndex = 5;
            this.btnOutput.Text = "数据导出";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.btnOutput_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(415, 511);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 21);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // sqlConn
            // 
            this.sqlConn.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlDA
            // 
            this.sqlDA.DeleteCommand = this.sqlDeleteCommand1;
            this.sqlDA.InsertCommand = this.sqlInsertCommand1;
            this.sqlDA.SelectCommand = this.sqlSelectCommand1;
            this.sqlDA.UpdateCommand = this.sqlUpdateCommand1;
            // 
            // openFileDialogoff
            // 
            this.openFileDialogoff.RestoreDirectory = true;
            // 
            // saveFileDialogOutput
            // 
            this.saveFileDialogOutput.CreatePrompt = true;
            this.saveFileDialogOutput.DefaultExt = "xls";
            this.saveFileDialogOutput.Filter = "excel文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            this.saveFileDialogOutput.RestoreDirectory = true;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Azure;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.labelConference);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(578, 49);
            this.panel1.TabIndex = 7;
            // 
            // labelConference
            // 
            this.labelConference.AutoSize = true;
            this.labelConference.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelConference.Location = new System.Drawing.Point(11, 18);
            this.labelConference.Name = "labelConference";
            this.labelConference.Size = new System.Drawing.Size(0, 19);
            this.labelConference.TabIndex = 0;
            // 
            // formCount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(578, 560);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOutput);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.statusStripCount);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formCount";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "会议统计";
            this.Load += new System.EventHandler(this.formCount_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.statusStripCount.ResumeLayout(false);
            this.statusStripCount.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSign)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewUnSign)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNew)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxInfo;
        private System.Windows.Forms.ListBox listBoxInput;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnUp;
        private System.Windows.Forms.Button btnDown;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBoxderysl;
        private System.Windows.Forms.CheckBox checkBoxqdrysl;
        private System.Windows.Forms.CheckBox checkBoxqxrysl;
        private System.Windows.Forms.CheckBox checkBoxcdrysl;
        private System.Windows.Forms.CheckBox checkBoxqcsj;
        private System.Windows.Forms.CheckBox checkBoxhztjxx;
        private System.Windows.Forms.CheckBox checkBoxcdrylb;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnOffInput;
        private System.Windows.Forms.StatusStrip statusStripCount;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelCount;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnCount;
        private System.Windows.Forms.Button btnOutput;
        private System.Windows.Forms.Button btnCancel;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlDataAdapter sqlDA;
        private System.Windows.Forms.OpenFileDialog openFileDialogoff;
        private System.Windows.Forms.SaveFileDialog saveFileDialogOutput;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelConference;
        private System.Windows.Forms.CheckBox checkBoxqxrylb;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dataGridViewSign;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dataGridViewUnSign;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dataGridViewNew;
        private System.Windows.Forms.CheckBox checkBoxNew;
        private System.Windows.Forms.Button btnSCJ;

    }
}