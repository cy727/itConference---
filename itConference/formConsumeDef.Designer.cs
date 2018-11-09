namespace itConference
{
    partial class formConsumeDef
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formConsumeDef));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelConference = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btncDel = new System.Windows.Forms.Button();
            this.btncEdit = new System.Windows.Forms.Button();
            this.btncAdd = new System.Windows.Forms.Button();
            this.gb1 = new System.Windows.Forms.GroupBox();
            this.dataGridViewConsume = new System.Windows.Forms.DataGridView();
            this.dateTimePickerCend = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePickerCstart = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxCname = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnCanCel = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDA = new System.Data.SqlClient.SqlDataAdapter();
            this.errorProviderC = new System.Windows.Forms.ErrorProvider(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.gb1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewConsume)).BeginInit();
            this.panel3.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderC)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.statusStrip1, 0, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(584, 390);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Azure;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.labelConference);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(578, 34);
            this.panel1.TabIndex = 0;
            // 
            // labelConference
            // 
            this.labelConference.AutoSize = true;
            this.labelConference.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelConference.Location = new System.Drawing.Point(6, 6);
            this.labelConference.Name = "labelConference";
            this.labelConference.Size = new System.Drawing.Size(75, 19);
            this.labelConference.TabIndex = 0;
            this.labelConference.Text = "label1";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btncDel);
            this.panel2.Controls.Add(this.btncEdit);
            this.panel2.Controls.Add(this.btncAdd);
            this.panel2.Controls.Add(this.gb1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 43);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(578, 284);
            this.panel2.TabIndex = 1;
            // 
            // btncDel
            // 
            this.btncDel.Location = new System.Drawing.Point(487, 198);
            this.btncDel.Name = "btncDel";
            this.btncDel.Size = new System.Drawing.Size(75, 23);
            this.btncDel.TabIndex = 3;
            this.btncDel.Text = "删除";
            this.btncDel.UseVisualStyleBackColor = true;
            this.btncDel.Click += new System.EventHandler(this.btncDel_Click);
            // 
            // btncEdit
            // 
            this.btncEdit.Location = new System.Drawing.Point(487, 127);
            this.btncEdit.Name = "btncEdit";
            this.btncEdit.Size = new System.Drawing.Size(75, 23);
            this.btncEdit.TabIndex = 2;
            this.btncEdit.Text = "修改";
            this.btncEdit.UseVisualStyleBackColor = true;
            this.btncEdit.Click += new System.EventHandler(this.btncEdit_Click);
            // 
            // btncAdd
            // 
            this.btncAdd.Location = new System.Drawing.Point(487, 63);
            this.btncAdd.Name = "btncAdd";
            this.btncAdd.Size = new System.Drawing.Size(75, 23);
            this.btncAdd.TabIndex = 1;
            this.btncAdd.Text = "添加";
            this.btncAdd.UseVisualStyleBackColor = true;
            this.btncAdd.Click += new System.EventHandler(this.btncAdd_Click);
            // 
            // gb1
            // 
            this.gb1.Controls.Add(this.dataGridViewConsume);
            this.gb1.Controls.Add(this.dateTimePickerCend);
            this.gb1.Controls.Add(this.label3);
            this.gb1.Controls.Add(this.dateTimePickerCstart);
            this.gb1.Controls.Add(this.label2);
            this.gb1.Controls.Add(this.textBoxCname);
            this.gb1.Controls.Add(this.label1);
            this.gb1.Location = new System.Drawing.Point(13, 4);
            this.gb1.Name = "gb1";
            this.gb1.Size = new System.Drawing.Size(458, 270);
            this.gb1.TabIndex = 0;
            this.gb1.TabStop = false;
            this.gb1.Text = "会议消费";
            // 
            // dataGridViewConsume
            // 
            this.dataGridViewConsume.AllowUserToAddRows = false;
            this.dataGridViewConsume.AllowUserToDeleteRows = false;
            this.dataGridViewConsume.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewConsume.Location = new System.Drawing.Point(7, 94);
            this.dataGridViewConsume.Name = "dataGridViewConsume";
            this.dataGridViewConsume.ReadOnly = true;
            this.dataGridViewConsume.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewConsume.Size = new System.Drawing.Size(436, 170);
            this.dataGridViewConsume.TabIndex = 6;
            this.dataGridViewConsume.SelectionChanged += new System.EventHandler(this.dataGridViewConsume_SelectionChanged);
            // 
            // dateTimePickerCend
            // 
            this.dateTimePickerCend.CustomFormat = "yyyy年MMMMdd日 dddd   HH:mm";
            this.dateTimePickerCend.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerCend.Location = new System.Drawing.Point(80, 68);
            this.dateTimePickerCend.Name = "dateTimePickerCend";
            this.dateTimePickerCend.Size = new System.Drawing.Size(363, 20);
            this.dateTimePickerCend.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 72);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "结束时间：";
            // 
            // dateTimePickerCstart
            // 
            this.dateTimePickerCstart.CustomFormat = "yyyy年MMMMdd日 dddd   HH:mm";
            this.dateTimePickerCstart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerCstart.Location = new System.Drawing.Point(80, 43);
            this.dateTimePickerCstart.Name = "dateTimePickerCstart";
            this.dateTimePickerCstart.Size = new System.Drawing.Size(363, 20);
            this.dateTimePickerCstart.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "开始时间：";
            // 
            // textBoxCname
            // 
            this.textBoxCname.Location = new System.Drawing.Point(80, 17);
            this.textBoxCname.Name = "textBoxCname";
            this.textBoxCname.Size = new System.Drawing.Size(363, 20);
            this.textBoxCname.TabIndex = 1;
            this.textBoxCname.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxCname_Validating);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "消费名称：";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.White;
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel3.Controls.Add(this.btnCanCel);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(3, 333);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(578, 34);
            this.panel3.TabIndex = 2;
            // 
            // btnCanCel
            // 
            this.btnCanCel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnCanCel.Location = new System.Drawing.Point(488, 4);
            this.btnCanCel.Name = "btnCanCel";
            this.btnCanCel.Size = new System.Drawing.Size(75, 23);
            this.btnCanCel.TabIndex = 0;
            this.btnCanCel.Text = "关闭";
            this.btnCanCel.UseVisualStyleBackColor = true;
            this.btnCanCel.Click += new System.EventHandler(this.btnCanCel_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 370);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(584, 20);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 15);
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
            // errorProviderC
            // 
            this.errorProviderC.ContainerControl = this;
            // 
            // formConsumeDef
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 390);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formConsumeDef";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "消费定义";
            this.Load += new System.EventHandler(this.formConsumeDef_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.gb1.ResumeLayout(false);
            this.gb1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewConsume)).EndInit();
            this.panel3.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderC)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelConference;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnCanCel;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.Button btncDel;
        private System.Windows.Forms.Button btncEdit;
        private System.Windows.Forms.Button btncAdd;
        private System.Windows.Forms.GroupBox gb1;
        private System.Windows.Forms.DateTimePicker dateTimePickerCend;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dateTimePickerCstart;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxCname;
        private System.Windows.Forms.Label label1;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlDataAdapter sqlDA;
        private System.Windows.Forms.DataGridView dataGridViewConsume;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ErrorProvider errorProviderC;
    }
}