namespace itConference
{
    partial class formConference
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formConference));
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnGEdit = new System.Windows.Forms.Button();
            this.btnGDel = new System.Windows.Forms.Button();
            this.textBoxGroup = new System.Windows.Forms.TextBox();
            this.btnGAdd = new System.Windows.Forms.Button();
            this.listBoxGroup = new System.Windows.Forms.ListBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dataGridViewPeople = new System.Windows.Forms.DataGridView();
            this.btnYes = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnPSelect = new System.Windows.Forms.Button();
            this.btnPSelectI = new System.Windows.Forms.Button();
            this.btnPSelectAll = new System.Windows.Forms.Button();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDA = new System.Data.SqlClient.SqlDataAdapter();
            this.btnGroup = new System.Windows.Forms.Button();
            this.statusStripConference = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelConference = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnGClear = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnGSelect = new System.Windows.Forms.Button();
            this.btnGUnselect = new System.Windows.Forms.Button();
            this.comboBoxGroup = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnSite = new System.Windows.Forms.Button();
            this.textBoxSite = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.errorProviderConference = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnInput = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.btnWCH = new System.Windows.Forms.Button();
            this.btnCH = new System.Windows.Forms.Button();
            this.btnS = new System.Windows.Forms.Button();
            this.btnsAll = new System.Windows.Forms.Button();
            this.comboBoxSZW = new System.Windows.Forms.ComboBox();
            this.textBoxXBH = new System.Windows.Forms.TextBox();
            this.comboBoxSDW = new System.Windows.Forms.ComboBox();
            this.textBoxSXM = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.openFileDialogExcel = new System.Windows.Forms.OpenFileDialog();
            this.btnTZ = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPeople)).BeginInit();
            this.statusStripConference.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderConference)).BeginInit();
            this.groupBox6.SuspendLayout();
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
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(553, 144);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "会议基本信息";
            // 
            // textBoxConferenceNo
            // 
            this.textBoxConferenceNo.Location = new System.Drawing.Point(423, 49);
            this.textBoxConferenceNo.Name = "textBoxConferenceNo";
            this.textBoxConferenceNo.Size = new System.Drawing.Size(118, 20);
            this.textBoxConferenceNo.TabIndex = 13;
            this.textBoxConferenceNo.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxConferenceNo_Validating);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(369, 52);
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
            this.textBoxCintro.Size = new System.Drawing.Size(471, 38);
            this.textBoxCintro.TabIndex = 5;
            this.textBoxCintro.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxCintro_Validating);
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
            this.textBoxCname2.Size = new System.Drawing.Size(293, 20);
            this.textBoxCname2.TabIndex = 3;
            this.textBoxCname2.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxCname2_Validating);
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
            this.textBoxCname.Size = new System.Drawing.Size(471, 20);
            this.textBoxCname.TabIndex = 1;
            this.textBoxCname.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxCname_Validating);
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
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnGEdit);
            this.groupBox2.Controls.Add(this.btnGDel);
            this.groupBox2.Controls.Add(this.textBoxGroup);
            this.groupBox2.Controls.Add(this.btnGAdd);
            this.groupBox2.Controls.Add(this.listBoxGroup);
            this.groupBox2.Location = new System.Drawing.Point(573, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(207, 144);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "会议人员分组";
            // 
            // btnGEdit
            // 
            this.btnGEdit.Location = new System.Drawing.Point(92, 118);
            this.btnGEdit.Name = "btnGEdit";
            this.btnGEdit.Size = new System.Drawing.Size(50, 23);
            this.btnGEdit.TabIndex = 4;
            this.btnGEdit.Text = "修改";
            this.btnGEdit.UseVisualStyleBackColor = true;
            this.btnGEdit.Click += new System.EventHandler(this.btnGEdit_Click);
            // 
            // btnGDel
            // 
            this.btnGDel.Location = new System.Drawing.Point(147, 118);
            this.btnGDel.Name = "btnGDel";
            this.btnGDel.Size = new System.Drawing.Size(50, 23);
            this.btnGDel.TabIndex = 3;
            this.btnGDel.Text = "删除";
            this.btnGDel.UseVisualStyleBackColor = true;
            this.btnGDel.Click += new System.EventHandler(this.btnGDel_Click);
            // 
            // textBoxGroup
            // 
            this.textBoxGroup.Location = new System.Drawing.Point(7, 19);
            this.textBoxGroup.Name = "textBoxGroup";
            this.textBoxGroup.Size = new System.Drawing.Size(191, 20);
            this.textBoxGroup.TabIndex = 2;
            this.textBoxGroup.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxGroup_Validating);
            // 
            // btnGAdd
            // 
            this.btnGAdd.Location = new System.Drawing.Point(39, 118);
            this.btnGAdd.Name = "btnGAdd";
            this.btnGAdd.Size = new System.Drawing.Size(50, 23);
            this.btnGAdd.TabIndex = 1;
            this.btnGAdd.Text = "增加";
            this.btnGAdd.UseVisualStyleBackColor = true;
            this.btnGAdd.Click += new System.EventHandler(this.btnGAdd_Click);
            // 
            // listBoxGroup
            // 
            this.listBoxGroup.FormattingEnabled = true;
            this.listBoxGroup.Location = new System.Drawing.Point(6, 45);
            this.listBoxGroup.Name = "listBoxGroup";
            this.listBoxGroup.Size = new System.Drawing.Size(192, 69);
            this.listBoxGroup.TabIndex = 0;
            this.listBoxGroup.SelectedIndexChanged += new System.EventHandler(this.listBoxGroup_SelectedIndexChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dataGridViewPeople);
            this.groupBox3.Location = new System.Drawing.Point(13, 297);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(770, 328);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "参会人员";
            // 
            // dataGridViewPeople
            // 
            this.dataGridViewPeople.AllowUserToAddRows = false;
            this.dataGridViewPeople.AllowUserToDeleteRows = false;
            this.dataGridViewPeople.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewPeople.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewPeople.Location = new System.Drawing.Point(3, 16);
            this.dataGridViewPeople.Name = "dataGridViewPeople";
            this.dataGridViewPeople.Size = new System.Drawing.Size(764, 309);
            this.dataGridViewPeople.TabIndex = 0;
            this.dataGridViewPeople.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellValueChanged);
            this.dataGridViewPeople.Click += new System.EventHandler(this.dataGridViewPeople_Click);
            this.dataGridViewPeople.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewPeople_CellContentClick);
            // 
            // btnYes
            // 
            this.btnYes.Location = new System.Drawing.Point(624, 631);
            this.btnYes.Name = "btnYes";
            this.btnYes.Size = new System.Drawing.Size(75, 23);
            this.btnYes.TabIndex = 3;
            this.btnYes.Text = "确定";
            this.btnYes.UseVisualStyleBackColor = true;
            this.btnYes.Click += new System.EventHandler(this.btnYes_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(703, 631);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnPSelect
            // 
            this.btnPSelect.Location = new System.Drawing.Point(20, 631);
            this.btnPSelect.Name = "btnPSelect";
            this.btnPSelect.Size = new System.Drawing.Size(66, 23);
            this.btnPSelect.TabIndex = 5;
            this.btnPSelect.Text = "标记高亮";
            this.btnPSelect.UseVisualStyleBackColor = true;
            this.btnPSelect.Click += new System.EventHandler(this.btnPSelect_Click);
            // 
            // btnPSelectI
            // 
            this.btnPSelectI.Location = new System.Drawing.Point(89, 631);
            this.btnPSelectI.Name = "btnPSelectI";
            this.btnPSelectI.Size = new System.Drawing.Size(60, 23);
            this.btnPSelectI.TabIndex = 6;
            this.btnPSelectI.Text = "反选";
            this.btnPSelectI.UseVisualStyleBackColor = true;
            this.btnPSelectI.Click += new System.EventHandler(this.btnPSelectI_Click);
            // 
            // btnPSelectAll
            // 
            this.btnPSelectAll.Location = new System.Drawing.Point(155, 631);
            this.btnPSelectAll.Name = "btnPSelectAll";
            this.btnPSelectAll.Size = new System.Drawing.Size(60, 23);
            this.btnPSelectAll.TabIndex = 7;
            this.btnPSelectAll.Text = "全选";
            this.btnPSelectAll.UseVisualStyleBackColor = true;
            this.btnPSelectAll.Click += new System.EventHandler(this.btnPSelectAll_Click);
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
            // btnGroup
            // 
            this.btnGroup.Location = new System.Drawing.Point(319, 631);
            this.btnGroup.Name = "btnGroup";
            this.btnGroup.Size = new System.Drawing.Size(75, 23);
            this.btnGroup.TabIndex = 8;
            this.btnGroup.Text = "人员分组";
            this.btnGroup.UseVisualStyleBackColor = true;
            this.btnGroup.Click += new System.EventHandler(this.btnGroup_Click);
            // 
            // statusStripConference
            // 
            this.statusStripConference.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelConference});
            this.statusStripConference.Location = new System.Drawing.Point(0, 662);
            this.statusStripConference.Name = "statusStripConference";
            this.statusStripConference.Size = new System.Drawing.Size(790, 22);
            this.statusStripConference.TabIndex = 9;
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
            this.btnClear.Location = new System.Drawing.Point(220, 631);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(56, 23);
            this.btnClear.TabIndex = 10;
            this.btnClear.Text = "清除";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnGClear
            // 
            this.btnGClear.Location = new System.Drawing.Point(400, 631);
            this.btnGClear.Name = "btnGClear";
            this.btnGClear.Size = new System.Drawing.Size(75, 23);
            this.btnGClear.TabIndex = 11;
            this.btnGClear.Text = "清除分组";
            this.btnGClear.UseVisualStyleBackColor = true;
            this.btnGClear.Click += new System.EventHandler(this.btnGClear_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnGSelect);
            this.groupBox4.Controls.Add(this.btnGUnselect);
            this.groupBox4.Controls.Add(this.comboBoxGroup);
            this.groupBox4.Controls.Add(this.label8);
            this.groupBox4.Location = new System.Drawing.Point(13, 162);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(440, 47);
            this.groupBox4.TabIndex = 12;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "人员分组选择";
            // 
            // btnGSelect
            // 
            this.btnGSelect.Location = new System.Drawing.Point(306, 15);
            this.btnGSelect.Name = "btnGSelect";
            this.btnGSelect.Size = new System.Drawing.Size(60, 23);
            this.btnGSelect.TabIndex = 3;
            this.btnGSelect.Text = "选择";
            this.btnGSelect.UseVisualStyleBackColor = true;
            this.btnGSelect.Click += new System.EventHandler(this.btnGSelect_Click);
            // 
            // btnGUnselect
            // 
            this.btnGUnselect.Location = new System.Drawing.Point(372, 15);
            this.btnGUnselect.Name = "btnGUnselect";
            this.btnGUnselect.Size = new System.Drawing.Size(60, 23);
            this.btnGUnselect.TabIndex = 2;
            this.btnGUnselect.Text = "反选";
            this.btnGUnselect.UseVisualStyleBackColor = true;
            this.btnGUnselect.Click += new System.EventHandler(this.btnGUnselect_Click);
            // 
            // comboBoxGroup
            // 
            this.comboBoxGroup.FormattingEnabled = true;
            this.comboBoxGroup.Location = new System.Drawing.Point(81, 17);
            this.comboBoxGroup.Name = "comboBoxGroup";
            this.comboBoxGroup.Size = new System.Drawing.Size(219, 21);
            this.comboBoxGroup.TabIndex = 1;
            this.comboBoxGroup.Validating += new System.ComponentModel.CancelEventHandler(this.comboBoxGroup_Validating);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(19, 20);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(67, 13);
            this.label8.TabIndex = 0;
            this.label8.Text = "人员类别：";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnSite);
            this.groupBox5.Controls.Add(this.textBoxSite);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Location = new System.Drawing.Point(459, 162);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(321, 47);
            this.groupBox5.TabIndex = 13;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "座位号";
            // 
            // btnSite
            // 
            this.btnSite.Location = new System.Drawing.Point(237, 15);
            this.btnSite.Name = "btnSite";
            this.btnSite.Size = new System.Drawing.Size(75, 23);
            this.btnSite.TabIndex = 2;
            this.btnSite.Text = "修改";
            this.btnSite.UseVisualStyleBackColor = true;
            this.btnSite.Click += new System.EventHandler(this.btnSite_Click);
            // 
            // textBoxSite
            // 
            this.textBoxSite.Location = new System.Drawing.Point(78, 17);
            this.textBoxSite.Name = "textBoxSite";
            this.textBoxSite.Size = new System.Drawing.Size(153, 20);
            this.textBoxSite.TabIndex = 1;
            this.textBoxSite.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxSite_Validating);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(55, 13);
            this.label7.TabIndex = 0;
            this.label7.Text = "座位号：";
            // 
            // errorProviderConference
            // 
            this.errorProviderConference.ContainerControl = this;
            // 
            // btnInput
            // 
            this.btnInput.Location = new System.Drawing.Point(517, 631);
            this.btnInput.Name = "btnInput";
            this.btnInput.Size = new System.Drawing.Size(75, 23);
            this.btnInput.TabIndex = 14;
            this.btnInput.Text = "会议导入";
            this.btnInput.UseVisualStyleBackColor = true;
            this.btnInput.Visible = false;
            this.btnInput.Click += new System.EventHandler(this.btnInput_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.btnWCH);
            this.groupBox6.Controls.Add(this.btnCH);
            this.groupBox6.Controls.Add(this.btnS);
            this.groupBox6.Controls.Add(this.btnsAll);
            this.groupBox6.Controls.Add(this.comboBoxSZW);
            this.groupBox6.Controls.Add(this.textBoxXBH);
            this.groupBox6.Controls.Add(this.comboBoxSDW);
            this.groupBox6.Controls.Add(this.textBoxSXM);
            this.groupBox6.Controls.Add(this.label12);
            this.groupBox6.Controls.Add(this.label11);
            this.groupBox6.Controls.Add(this.label10);
            this.groupBox6.Controls.Add(this.label9);
            this.groupBox6.Location = new System.Drawing.Point(13, 216);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(767, 75);
            this.groupBox6.TabIndex = 15;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "人员过滤";
            // 
            // btnWCH
            // 
            this.btnWCH.Location = new System.Drawing.Point(560, 40);
            this.btnWCH.Name = "btnWCH";
            this.btnWCH.Size = new System.Drawing.Size(89, 23);
            this.btnWCH.TabIndex = 13;
            this.btnWCH.Text = "未参会人员";
            this.btnWCH.UseVisualStyleBackColor = true;
            this.btnWCH.Click += new System.EventHandler(this.btnWCH_Click);
            // 
            // btnCH
            // 
            this.btnCH.Location = new System.Drawing.Point(472, 40);
            this.btnCH.Name = "btnCH";
            this.btnCH.Size = new System.Drawing.Size(82, 23);
            this.btnCH.TabIndex = 12;
            this.btnCH.Text = "参会人员";
            this.btnCH.UseVisualStyleBackColor = true;
            this.btnCH.Click += new System.EventHandler(this.btnCH_Click);
            // 
            // btnS
            // 
            this.btnS.Location = new System.Drawing.Point(671, 40);
            this.btnS.Name = "btnS";
            this.btnS.Size = new System.Drawing.Size(75, 23);
            this.btnS.TabIndex = 11;
            this.btnS.Text = "筛选";
            this.btnS.UseVisualStyleBackColor = true;
            this.btnS.Click += new System.EventHandler(this.btnS_Click);
            // 
            // btnsAll
            // 
            this.btnsAll.Location = new System.Drawing.Point(671, 14);
            this.btnsAll.Name = "btnsAll";
            this.btnsAll.Size = new System.Drawing.Size(75, 23);
            this.btnsAll.TabIndex = 8;
            this.btnsAll.Text = "全部";
            this.btnsAll.UseVisualStyleBackColor = true;
            this.btnsAll.Click += new System.EventHandler(this.btnsAll_Click);
            // 
            // comboBoxSZW
            // 
            this.comboBoxSZW.FormattingEnabled = true;
            this.comboBoxSZW.Location = new System.Drawing.Point(250, 42);
            this.comboBoxSZW.Name = "comboBoxSZW";
            this.comboBoxSZW.Size = new System.Drawing.Size(210, 21);
            this.comboBoxSZW.TabIndex = 7;
            // 
            // textBoxXBH
            // 
            this.textBoxXBH.Location = new System.Drawing.Point(70, 42);
            this.textBoxXBH.Name = "textBoxXBH";
            this.textBoxXBH.Size = new System.Drawing.Size(125, 20);
            this.textBoxXBH.TabIndex = 6;
            // 
            // comboBoxSDW
            // 
            this.comboBoxSDW.FormattingEnabled = true;
            this.comboBoxSDW.Location = new System.Drawing.Point(250, 15);
            this.comboBoxSDW.Name = "comboBoxSDW";
            this.comboBoxSDW.Size = new System.Drawing.Size(399, 21);
            this.comboBoxSDW.TabIndex = 5;
            // 
            // textBoxSXM
            // 
            this.textBoxSXM.Location = new System.Drawing.Point(70, 17);
            this.textBoxSXM.Name = "textBoxSXM";
            this.textBoxSXM.Size = new System.Drawing.Size(125, 20);
            this.textBoxSXM.TabIndex = 4;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(22, 45);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(43, 13);
            this.label12.TabIndex = 3;
            this.label12.Text = "编号：";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(201, 45);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(43, 13);
            this.label11.TabIndex = 2;
            this.label11.Text = "职务：";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(201, 20);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(43, 13);
            this.label10.TabIndex = 1;
            this.label10.Text = "单位：";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(22, 20);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(43, 13);
            this.label9.TabIndex = 0;
            this.label9.Text = "姓名：";
            // 
            // openFileDialogExcel
            // 
            this.openFileDialogExcel.Filter = "excel文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            this.openFileDialogExcel.RestoreDirectory = true;
            // 
            // btnTZ
            // 
            this.btnTZ.Location = new System.Drawing.Point(517, 631);
            this.btnTZ.Name = "btnTZ";
            this.btnTZ.Size = new System.Drawing.Size(75, 23);
            this.btnTZ.TabIndex = 16;
            this.btnTZ.Text = "会议通知";
            this.btnTZ.UseVisualStyleBackColor = true;
            this.btnTZ.Click += new System.EventHandler(this.btnTZ_Click);
            // 
            // formConference
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(790, 684);
            this.Controls.Add(this.btnTZ);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.btnInput);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.btnGClear);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.statusStripConference);
            this.Controls.Add(this.btnGroup);
            this.Controls.Add(this.btnPSelectAll);
            this.Controls.Add(this.btnPSelectI);
            this.Controls.Add(this.btnPSelect);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnYes);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formConference";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "会议管理";
            this.Load += new System.EventHandler(this.formConference_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPeople)).EndInit();
            this.statusStripConference.ResumeLayout(false);
            this.statusStripConference.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderConference)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnYes;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnGAdd;
        private System.Windows.Forms.ListBox listBoxGroup;
        private System.Windows.Forms.Button btnGEdit;
        private System.Windows.Forms.Button btnGDel;
        private System.Windows.Forms.TextBox textBoxGroup;
        private System.Windows.Forms.Button btnPSelect;
        private System.Windows.Forms.Button btnPSelectI;
        private System.Windows.Forms.Button btnPSelectAll;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlDataAdapter sqlDA;
        private System.Windows.Forms.Button btnGroup;
        private System.Windows.Forms.StatusStrip statusStripConference;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelConference;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnGClear;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button btnSite;
        private System.Windows.Forms.TextBox textBoxSite;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnGSelect;
        private System.Windows.Forms.Button btnGUnselect;
        private System.Windows.Forms.ComboBox comboBoxGroup;
        private System.Windows.Forms.ErrorProvider errorProviderConference;
        private System.Windows.Forms.Button btnInput;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBoxSDW;
        private System.Windows.Forms.TextBox textBoxSXM;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button btnsAll;
        private System.Windows.Forms.ComboBox comboBoxSZW;
        private System.Windows.Forms.TextBox textBoxXBH;
        private System.Windows.Forms.Button btnS;
        private System.Windows.Forms.Button btnCH;
        private System.Windows.Forms.Button btnWCH;
        private System.Windows.Forms.OpenFileDialog openFileDialogExcel;
        private System.Windows.Forms.Button btnTZ;
        public System.Windows.Forms.TextBox textBoxCintro;
        public System.Windows.Forms.TextBox textBoxCname2;
        public System.Windows.Forms.TextBox textBoxCname;
        public System.Windows.Forms.DateTimePicker dateTimePickerStart;
        public System.Windows.Forms.DataGridView dataGridViewPeople;
        public System.Windows.Forms.TextBox textBoxConferenceNo;
    }
}