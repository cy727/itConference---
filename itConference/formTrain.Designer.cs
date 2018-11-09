namespace itConference
{
    partial class formTrain
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formTrain));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxConferenceNo = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.dateTimePickerStart = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxCintro = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxCname2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxCname = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.statusStripConference = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelConference = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnClear = new System.Windows.Forms.Button();
            this.errorProviderConference = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnWCH = new System.Windows.Forms.Button();
            this.btnS = new System.Windows.Forms.Button();
            this.btnCH = new System.Windows.Forms.Button();
            this.btnsAll = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.comboBoxSDW = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.btnPSelectAll = new System.Windows.Forms.Button();
            this.btnPSelect = new System.Windows.Forms.Button();
            this.btnPSelectI = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.btnYes = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dataGridViewPeople = new System.Windows.Forms.DataGridView();
            this.label7 = new System.Windows.Forms.Label();
            this.btnSite = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.numericUpDownCHRS = new System.Windows.Forms.NumericUpDown();
            this.btnInput = new System.Windows.Forms.Button();
            this.openFileDialogExcel = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.statusStripConference.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderConference)).BeginInit();
            this.groupBox6.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPeople)).BeginInit();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownCHRS)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxConferenceNo);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.dateTimePickerEnd);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.dateTimePickerStart);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.textBoxCintro);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxCname2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBoxCname);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(13, 7);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(770, 144);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "会议基本信息";
            // 
            // textBoxConferenceNo
            // 
            this.textBoxConferenceNo.Location = new System.Drawing.Point(627, 49);
            this.textBoxConferenceNo.Name = "textBoxConferenceNo";
            this.textBoxConferenceNo.Size = new System.Drawing.Size(137, 20);
            this.textBoxConferenceNo.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(557, 52);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "会场号：";
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.CustomFormat = "yyyy年MMMMdd日 dddd HH:mm";
            this.dateTimePickerEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerEnd.Location = new System.Drawing.Point(341, 117);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerEnd.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(278, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(67, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "结束时间：";
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.CustomFormat = "yyyy年MMMMdd日 dddd HH:mm";
            this.dateTimePickerStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerStart.Location = new System.Drawing.Point(70, 117);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(202, 20);
            this.dateTimePickerStart.TabIndex = 9;
            this.dateTimePickerStart.ValueChanged += new System.EventHandler(this.dateTimePickerStart_ValueChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "开始时间：";
            // 
            // textBoxCintro
            // 
            this.textBoxCintro.Location = new System.Drawing.Point(70, 74);
            this.textBoxCintro.Multiline = true;
            this.textBoxCintro.Name = "textBoxCintro";
            this.textBoxCintro.Size = new System.Drawing.Size(694, 38);
            this.textBoxCintro.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "会议简介：";
            // 
            // textBoxCname2
            // 
            this.textBoxCname2.Location = new System.Drawing.Point(70, 49);
            this.textBoxCname2.Name = "textBoxCname2";
            this.textBoxCname2.Size = new System.Drawing.Size(484, 20);
            this.textBoxCname2.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "会议副题：";
            // 
            // textBoxCname
            // 
            this.textBoxCname.Location = new System.Drawing.Point(70, 25);
            this.textBoxCname.Name = "textBoxCname";
            this.textBoxCname.Size = new System.Drawing.Size(694, 20);
            this.textBoxCname.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "会议名称：";
            // 
            // statusStripConference
            // 
            this.statusStripConference.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelConference});
            this.statusStripConference.Location = new System.Drawing.Point(0, 664);
            this.statusStripConference.Name = "statusStripConference";
            this.statusStripConference.Size = new System.Drawing.Size(792, 22);
            this.statusStripConference.TabIndex = 25;
            this.statusStripConference.Text = "statusStrip1";
            // 
            // toolStripStatusLabelConference
            // 
            this.toolStripStatusLabelConference.Name = "toolStripStatusLabelConference";
            this.toolStripStatusLabelConference.Size = new System.Drawing.Size(131, 17);
            this.toolStripStatusLabelConference.Text = "toolStripStatusLabel1";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(263, 625);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 26;
            this.btnClear.Text = "清除";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // errorProviderConference
            // 
            this.errorProviderConference.ContainerControl = this;
            // 
            // btnWCH
            // 
            this.btnWCH.Location = new System.Drawing.Point(406, 40);
            this.btnWCH.Name = "btnWCH";
            this.btnWCH.Size = new System.Drawing.Size(83, 23);
            this.btnWCH.TabIndex = 13;
            this.btnWCH.Text = "未参会人员";
            this.btnWCH.UseVisualStyleBackColor = true;
            this.btnWCH.Click += new System.EventHandler(this.btnWCH_Click);
            // 
            // btnS
            // 
            this.btnS.Location = new System.Drawing.Point(316, 11);
            this.btnS.Name = "btnS";
            this.btnS.Size = new System.Drawing.Size(82, 23);
            this.btnS.TabIndex = 11;
            this.btnS.Text = "筛选";
            this.btnS.UseVisualStyleBackColor = true;
            this.btnS.Click += new System.EventHandler(this.btnS_Click);
            // 
            // btnCH
            // 
            this.btnCH.Location = new System.Drawing.Point(316, 40);
            this.btnCH.Name = "btnCH";
            this.btnCH.Size = new System.Drawing.Size(82, 23);
            this.btnCH.TabIndex = 12;
            this.btnCH.Text = "参会人员";
            this.btnCH.UseVisualStyleBackColor = true;
            this.btnCH.Click += new System.EventHandler(this.btnCH_Click);
            // 
            // btnsAll
            // 
            this.btnsAll.Location = new System.Drawing.Point(406, 11);
            this.btnsAll.Name = "btnsAll";
            this.btnsAll.Size = new System.Drawing.Size(83, 23);
            this.btnsAll.TabIndex = 8;
            this.btnsAll.Text = "全部";
            this.btnsAll.UseVisualStyleBackColor = true;
            this.btnsAll.Click += new System.EventHandler(this.btnsAll_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.btnWCH);
            this.groupBox6.Controls.Add(this.btnCH);
            this.groupBox6.Controls.Add(this.btnS);
            this.groupBox6.Controls.Add(this.btnsAll);
            this.groupBox6.Controls.Add(this.comboBoxSDW);
            this.groupBox6.Controls.Add(this.label10);
            this.groupBox6.Location = new System.Drawing.Point(12, 157);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(495, 70);
            this.groupBox6.TabIndex = 31;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "过滤";
            // 
            // comboBoxSDW
            // 
            this.comboBoxSDW.FormattingEnabled = true;
            this.comboBoxSDW.Location = new System.Drawing.Point(25, 35);
            this.comboBoxSDW.Name = "comboBoxSDW";
            this.comboBoxSDW.Size = new System.Drawing.Size(285, 21);
            this.comboBoxSDW.TabIndex = 5;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(22, 19);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(43, 13);
            this.label10.TabIndex = 1;
            this.label10.Text = "单位：";
            // 
            // btnPSelectAll
            // 
            this.btnPSelectAll.Location = new System.Drawing.Point(182, 625);
            this.btnPSelectAll.Name = "btnPSelectAll";
            this.btnPSelectAll.Size = new System.Drawing.Size(75, 23);
            this.btnPSelectAll.TabIndex = 23;
            this.btnPSelectAll.Text = "全选";
            this.btnPSelectAll.UseVisualStyleBackColor = true;
            this.btnPSelectAll.Click += new System.EventHandler(this.btnPSelectAll_Click);
            // 
            // btnPSelect
            // 
            this.btnPSelect.Location = new System.Drawing.Point(20, 625);
            this.btnPSelect.Name = "btnPSelect";
            this.btnPSelect.Size = new System.Drawing.Size(75, 23);
            this.btnPSelect.TabIndex = 21;
            this.btnPSelect.Text = "标记高亮";
            this.btnPSelect.UseVisualStyleBackColor = true;
            this.btnPSelect.Click += new System.EventHandler(this.btnPSelect_Click);
            // 
            // btnPSelectI
            // 
            this.btnPSelectI.Location = new System.Drawing.Point(101, 625);
            this.btnPSelectI.Name = "btnPSelectI";
            this.btnPSelectI.Size = new System.Drawing.Size(75, 23);
            this.btnPSelectI.TabIndex = 22;
            this.btnPSelectI.Text = "反选";
            this.btnPSelectI.UseVisualStyleBackColor = true;
            this.btnPSelectI.Click += new System.EventHandler(this.btnPSelectI_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(702, 625);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 20;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnYes
            // 
            this.btnYes.Location = new System.Drawing.Point(621, 625);
            this.btnYes.Name = "btnYes";
            this.btnYes.Size = new System.Drawing.Size(75, 23);
            this.btnYes.TabIndex = 19;
            this.btnYes.Text = "确定";
            this.btnYes.UseVisualStyleBackColor = true;
            this.btnYes.Click += new System.EventHandler(this.btnYes_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dataGridViewPeople);
            this.groupBox3.Location = new System.Drawing.Point(13, 233);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(770, 386);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "参会单位";
            // 
            // dataGridViewPeople
            // 
            this.dataGridViewPeople.AllowUserToAddRows = false;
            this.dataGridViewPeople.AllowUserToDeleteRows = false;
            this.dataGridViewPeople.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewPeople.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewPeople.Location = new System.Drawing.Point(3, 16);
            this.dataGridViewPeople.Name = "dataGridViewPeople";
            this.dataGridViewPeople.Size = new System.Drawing.Size(764, 367);
            this.dataGridViewPeople.TabIndex = 0;
            this.dataGridViewPeople.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(67, 13);
            this.label7.TabIndex = 0;
            this.label7.Text = "参会人数：";
            // 
            // btnSite
            // 
            this.btnSite.Location = new System.Drawing.Point(177, 33);
            this.btnSite.Name = "btnSite";
            this.btnSite.Size = new System.Drawing.Size(75, 23);
            this.btnSite.TabIndex = 2;
            this.btnSite.Text = "修改";
            this.btnSite.UseVisualStyleBackColor = true;
            this.btnSite.Click += new System.EventHandler(this.btnSite_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.numericUpDownCHRS);
            this.groupBox5.Controls.Add(this.btnSite);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Location = new System.Drawing.Point(513, 157);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(271, 70);
            this.groupBox5.TabIndex = 29;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "参会人数";
            // 
            // numericUpDownCHRS
            // 
            this.numericUpDownCHRS.Location = new System.Drawing.Point(41, 36);
            this.numericUpDownCHRS.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numericUpDownCHRS.Name = "numericUpDownCHRS";
            this.numericUpDownCHRS.Size = new System.Drawing.Size(130, 20);
            this.numericUpDownCHRS.TabIndex = 3;
            this.numericUpDownCHRS.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnInput
            // 
            this.btnInput.Location = new System.Drawing.Point(492, 625);
            this.btnInput.Name = "btnInput";
            this.btnInput.Size = new System.Drawing.Size(75, 23);
            this.btnInput.TabIndex = 32;
            this.btnInput.Text = "会议导入";
            this.btnInput.UseVisualStyleBackColor = true;
            this.btnInput.Click += new System.EventHandler(this.btnInput_Click);
            // 
            // openFileDialogExcel
            // 
            this.openFileDialogExcel.Filter = "excel文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            this.openFileDialogExcel.RestoreDirectory = true;
            // 
            // formTrain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(792, 686);
            this.Controls.Add(this.btnInput);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStripConference);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.btnPSelectAll);
            this.Controls.Add(this.btnPSelect);
            this.Controls.Add(this.btnPSelectI);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnYes);
            this.Controls.Add(this.groupBox3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formTrain";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "培训管理";
            this.Load += new System.EventHandler(this.formTrain_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.statusStripConference.ResumeLayout(false);
            this.statusStripConference.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderConference)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPeople)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownCHRS)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBoxConferenceNo;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker dateTimePickerStart;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxCintro;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxCname2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxCname;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.StatusStrip statusStripConference;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelConference;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ErrorProvider errorProviderConference;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.NumericUpDown numericUpDownCHRS;
        private System.Windows.Forms.Button btnSite;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Button btnWCH;
        private System.Windows.Forms.Button btnCH;
        private System.Windows.Forms.Button btnS;
        private System.Windows.Forms.Button btnsAll;
        private System.Windows.Forms.ComboBox comboBoxSDW;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnPSelectAll;
        private System.Windows.Forms.Button btnPSelect;
        private System.Windows.Forms.Button btnPSelectI;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnYes;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.DataGridView dataGridViewPeople;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Windows.Forms.Button btnInput;
        private System.Windows.Forms.OpenFileDialog openFileDialogExcel;
    }
}