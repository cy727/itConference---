namespace itConference
{
    partial class formTZ
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formTZ));
            this.btnDOWN = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboBoxCOM = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.textBoxTZ = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.radioButtonSY = new System.Windows.Forms.RadioButton();
            this.radioButtonWTZ = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.textBoxRZ = new System.Windows.Forms.TextBox();
            this.btnRZ = new System.Windows.Forms.Button();
            this.saveFileDialogIO = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialogIO = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBarCount = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabelCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDOWN
            // 
            this.btnDOWN.Location = new System.Drawing.Point(598, 156);
            this.btnDOWN.Name = "btnDOWN";
            this.btnDOWN.Size = new System.Drawing.Size(75, 23);
            this.btnDOWN.TabIndex = 10;
            this.btnDOWN.Text = "发送通知";
            this.btnDOWN.UseVisualStyleBackColor = true;
            this.btnDOWN.Click += new System.EventHandler(this.btnDOWN_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(598, 206);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "关闭";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comboBoxCOM);
            this.groupBox1.Location = new System.Drawing.Point(487, 153);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(105, 73);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "端口值";
            // 
            // comboBoxCOM
            // 
            this.comboBoxCOM.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCOM.FormattingEnabled = true;
            this.comboBoxCOM.Items.AddRange(new object[] {
            "COM0",
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8"});
            this.comboBoxCOM.Location = new System.Drawing.Point(6, 30);
            this.comboBoxCOM.Name = "comboBoxCOM";
            this.comboBoxCOM.Size = new System.Drawing.Size(96, 21);
            this.comboBoxCOM.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBoxTZ);
            this.groupBox2.Location = new System.Drawing.Point(364, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(312, 138);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "通知内容";
            // 
            // textBoxTZ
            // 
            this.textBoxTZ.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxTZ.Location = new System.Drawing.Point(3, 16);
            this.textBoxTZ.Multiline = true;
            this.textBoxTZ.Name = "textBoxTZ";
            this.textBoxTZ.Size = new System.Drawing.Size(306, 119);
            this.textBoxTZ.TabIndex = 0;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.radioButtonWTZ);
            this.groupBox3.Controls.Add(this.radioButtonSY);
            this.groupBox3.Location = new System.Drawing.Point(364, 153);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(120, 73);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "范围";
            // 
            // radioButtonSY
            // 
            this.radioButtonSY.AutoSize = true;
            this.radioButtonSY.Location = new System.Drawing.Point(7, 20);
            this.radioButtonSY.Name = "radioButtonSY";
            this.radioButtonSY.Size = new System.Drawing.Size(97, 17);
            this.radioButtonSY.TabIndex = 0;
            this.radioButtonSY.Text = "所有参会人员";
            this.radioButtonSY.UseVisualStyleBackColor = true;
            // 
            // radioButtonWTZ
            // 
            this.radioButtonWTZ.AutoSize = true;
            this.radioButtonWTZ.Checked = true;
            this.radioButtonWTZ.Location = new System.Drawing.Point(7, 43);
            this.radioButtonWTZ.Name = "radioButtonWTZ";
            this.radioButtonWTZ.Size = new System.Drawing.Size(109, 17);
            this.radioButtonWTZ.TabIndex = 0;
            this.radioButtonWTZ.TabStop = true;
            this.radioButtonWTZ.Text = "参会未通知人员";
            this.radioButtonWTZ.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.textBoxRZ);
            this.groupBox4.Location = new System.Drawing.Point(12, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(346, 214);
            this.groupBox4.TabIndex = 13;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "发送日志";
            // 
            // textBoxRZ
            // 
            this.textBoxRZ.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxRZ.Location = new System.Drawing.Point(3, 16);
            this.textBoxRZ.Multiline = true;
            this.textBoxRZ.Name = "textBoxRZ";
            this.textBoxRZ.Size = new System.Drawing.Size(340, 195);
            this.textBoxRZ.TabIndex = 0;
            // 
            // btnRZ
            // 
            this.btnRZ.Location = new System.Drawing.Point(598, 181);
            this.btnRZ.Name = "btnRZ";
            this.btnRZ.Size = new System.Drawing.Size(75, 23);
            this.btnRZ.TabIndex = 14;
            this.btnRZ.Text = "日志保存";
            this.btnRZ.UseVisualStyleBackColor = true;
            this.btnRZ.Click += new System.EventHandler(this.btnRZ_Click);
            // 
            // saveFileDialogIO
            // 
            this.saveFileDialogIO.Filter = "2007 EXCEL files(*.xlsx)|*.xlsx|EXCEL files(*.xls)|*.xls";
            this.saveFileDialogIO.RestoreDirectory = true;
            // 
            // openFileDialogIO
            // 
            this.openFileDialogIO.Filter = "2007 EXCEL files(*.xlsx)|*.xlsx|EXCEL files(*.xls)|*.xls";
            this.openFileDialogIO.RestoreDirectory = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelStatus,
            this.toolStripProgressBarCount,
            this.toolStripStatusLabelCount});
            this.statusStrip1.Location = new System.Drawing.Point(0, 238);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(685, 22);
            this.statusStrip1.TabIndex = 15;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabelStatus
            // 
            this.toolStripStatusLabelStatus.Name = "toolStripStatusLabelStatus";
            this.toolStripStatusLabelStatus.Size = new System.Drawing.Size(445, 17);
            this.toolStripStatusLabelStatus.Spring = true;
            this.toolStripStatusLabelStatus.Text = "发送状态";
            this.toolStripStatusLabelStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // toolStripProgressBarCount
            // 
            this.toolStripProgressBarCount.Name = "toolStripProgressBarCount";
            this.toolStripProgressBarCount.Size = new System.Drawing.Size(200, 16);
            // 
            // toolStripStatusLabelCount
            // 
            this.toolStripStatusLabelCount.Name = "toolStripStatusLabelCount";
            this.toolStripStatusLabelCount.Size = new System.Drawing.Size(23, 17);
            this.toolStripStatusLabelCount.Text = "0/0";
            // 
            // formTZ
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(685, 260);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.btnRZ);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnDOWN);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formTZ";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "会议通知";
            this.Load += new System.EventHandler(this.formTZ_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDOWN;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.ComboBox comboBoxCOM;
        private System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.TextBox textBoxTZ;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton radioButtonWTZ;
        private System.Windows.Forms.RadioButton radioButtonSY;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox textBoxRZ;
        private System.Windows.Forms.Button btnRZ;
        private System.Windows.Forms.SaveFileDialog saveFileDialogIO;
        private System.Windows.Forms.OpenFileDialog openFileDialogIO;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelStatus;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBarCount;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelCount;
    }
}