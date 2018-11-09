namespace itConference
{
    partial class formoffSign
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formoffSign));
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelWarn = new System.Windows.Forms.ToolStripStatusLabel();
            this.gb1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBoxSign = new System.Windows.Forms.TextBox();
            this.dataGridViewSign = new System.Windows.Forms.DataGridView();
            this.btnRead = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.textBoxConference = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.saveFileDialogSign = new System.Windows.Forms.SaveFileDialog();
            this.errorProviderSign = new System.Windows.Forms.ErrorProvider(this.components);
            this.timerS = new System.Windows.Forms.Timer(this.components);
            this.buttonEXCEL = new System.Windows.Forms.Button();
            this.saveFileDialogExcel = new System.Windows.Forms.SaveFileDialog();
            this.statusStrip1.SuspendLayout();
            this.gb1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSign)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderSign)).BeginInit();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelCount,
            this.toolStripStatusLabel1,
            this.toolStripStatusLabelWarn});
            this.statusStrip1.Location = new System.Drawing.Point(0, 444);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(792, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabelCount
            // 
            this.toolStripStatusLabelCount.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left;
            this.toolStripStatusLabelCount.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStripStatusLabelCount.Name = "toolStripStatusLabelCount";
            this.toolStripStatusLabelCount.Size = new System.Drawing.Size(4, 17);
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(773, 17);
            this.toolStripStatusLabel1.Spring = true;
            // 
            // toolStripStatusLabelWarn
            // 
            this.toolStripStatusLabelWarn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStripStatusLabelWarn.ForeColor = System.Drawing.Color.Red;
            this.toolStripStatusLabelWarn.Name = "toolStripStatusLabelWarn";
            this.toolStripStatusLabelWarn.Size = new System.Drawing.Size(0, 17);
            // 
            // gb1
            // 
            this.gb1.Controls.Add(this.panel1);
            this.gb1.Controls.Add(this.dataGridViewSign);
            this.gb1.Controls.Add(this.textBoxSign);
            this.gb1.Location = new System.Drawing.Point(12, 64);
            this.gb1.Name = "gb1";
            this.gb1.Size = new System.Drawing.Size(768, 348);
            this.gb1.TabIndex = 1;
            this.gb1.TabStop = false;
            this.gb1.Text = "签到列表";
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(720, 352);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(205, 100);
            this.panel1.TabIndex = 7;
            // 
            // textBoxSign
            // 
            this.textBoxSign.Location = new System.Drawing.Point(652, 256);
            this.textBoxSign.Name = "textBoxSign";
            this.textBoxSign.Size = new System.Drawing.Size(85, 20);
            this.textBoxSign.TabIndex = 6;
            this.textBoxSign.TextChanged += new System.EventHandler(this.textBoxSign_TextChanged);
            // 
            // dataGridViewSign
            // 
            this.dataGridViewSign.AllowUserToAddRows = false;
            this.dataGridViewSign.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSign.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewSign.Location = new System.Drawing.Point(3, 16);
            this.dataGridViewSign.Name = "dataGridViewSign";
            this.dataGridViewSign.Size = new System.Drawing.Size(762, 329);
            this.dataGridViewSign.TabIndex = 0;
            // 
            // btnRead
            // 
            this.btnRead.Location = new System.Drawing.Point(618, 418);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(75, 23);
            this.btnRead.TabIndex = 3;
            this.btnRead.Text = "签到";
            this.btnRead.UseVisualStyleBackColor = true;
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(699, 418);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBoxConference);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(10, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(770, 45);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "会议信息";
            // 
            // textBoxConference
            // 
            this.textBoxConference.Location = new System.Drawing.Point(58, 17);
            this.textBoxConference.Name = "textBoxConference";
            this.textBoxConference.Size = new System.Drawing.Size(706, 20);
            this.textBoxConference.TabIndex = 1;
            this.textBoxConference.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxConference_Validating);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "会议名：";
            // 
            // saveFileDialogSign
            // 
            this.saveFileDialogSign.CreatePrompt = true;
            this.saveFileDialogSign.DefaultExt = "xml";
            this.saveFileDialogSign.Filter = "XML 文件(*.xml)|*.xml";
            this.saveFileDialogSign.RestoreDirectory = true;
            // 
            // errorProviderSign
            // 
            this.errorProviderSign.ContainerControl = this;
            // 
            // timerS
            // 
            this.timerS.Interval = 500;
            this.timerS.Tick += new System.EventHandler(this.timerS_Tick);
            // 
            // buttonEXCEL
            // 
            this.buttonEXCEL.Location = new System.Drawing.Point(22, 418);
            this.buttonEXCEL.Name = "buttonEXCEL";
            this.buttonEXCEL.Size = new System.Drawing.Size(100, 23);
            this.buttonEXCEL.TabIndex = 6;
            this.buttonEXCEL.Text = "输出到EXCEL";
            this.buttonEXCEL.UseVisualStyleBackColor = true;
            this.buttonEXCEL.Click += new System.EventHandler(this.buttonEXCEL_Click);
            // 
            // saveFileDialogExcel
            // 
            this.saveFileDialogExcel.CreatePrompt = true;
            this.saveFileDialogExcel.DefaultExt = "xls";
            this.saveFileDialogExcel.Filter = "excel文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            this.saveFileDialogExcel.RestoreDirectory = true;
            // 
            // formoffSign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(792, 466);
            this.Controls.Add(this.buttonEXCEL);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRead);
            this.Controls.Add(this.gb1);
            this.Controls.Add(this.statusStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formoffSign";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "离线签到";
            this.Load += new System.EventHandler(this.formoffSign_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.formoffSign_FormClosed);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.gb1.ResumeLayout(false);
            this.gb1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSign)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderSign)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.GroupBox gb1;
        private System.Windows.Forms.Button btnRead;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridView dataGridViewSign;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox textBoxConference;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.SaveFileDialog saveFileDialogSign;
        private System.Windows.Forms.TextBox textBoxSign;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ErrorProvider errorProviderSign;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelCount;
        private System.Windows.Forms.Timer timerS;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelWarn;
        private System.Windows.Forms.Button buttonEXCEL;
        private System.Windows.Forms.SaveFileDialog saveFileDialogExcel;
    }
}