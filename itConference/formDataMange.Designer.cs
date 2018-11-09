namespace itConference
{
    partial class formDataMange
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formDataMange));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnPeopleClear = new System.Windows.Forms.Button();
            this.btnConClear = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.textBoxLJHF = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRestore = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnBKPath = new System.Windows.Forms.Button();
            this.textBoxLJBK = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnBackUp = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.saveFileDialogData = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialogData = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnPeopleClear);
            this.groupBox1.Controls.Add(this.btnConClear);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(315, 114);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据清理";
            // 
            // btnPeopleClear
            // 
            this.btnPeopleClear.Location = new System.Drawing.Point(82, 71);
            this.btnPeopleClear.Name = "btnPeopleClear";
            this.btnPeopleClear.Size = new System.Drawing.Size(162, 23);
            this.btnPeopleClear.TabIndex = 1;
            this.btnPeopleClear.Text = "清空人员数据";
            this.btnPeopleClear.UseVisualStyleBackColor = true;
            this.btnPeopleClear.Click += new System.EventHandler(this.btnPeopleClear_Click);
            // 
            // btnConClear
            // 
            this.btnConClear.Location = new System.Drawing.Point(82, 33);
            this.btnConClear.Name = "btnConClear";
            this.btnConClear.Size = new System.Drawing.Size(162, 23);
            this.btnConClear.TabIndex = 0;
            this.btnConClear.Text = "清空会议数据";
            this.btnConClear.UseVisualStyleBackColor = true;
            this.btnConClear.Click += new System.EventHandler(this.btnConClear_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(133, 316);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "关闭";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // sqlConn
            // 
            this.sqlConn.FireInfoMessageEventOnUserErrors = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button1);
            this.groupBox3.Controls.Add(this.textBoxLJHF);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.btnRestore);
            this.groupBox3.Location = new System.Drawing.Point(13, 223);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(315, 87);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "数据库恢复";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(281, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(28, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBoxLJHF
            // 
            this.textBoxLJHF.Location = new System.Drawing.Point(70, 25);
            this.textBoxLJHF.Name = "textBoxLJHF";
            this.textBoxLJHF.Size = new System.Drawing.Size(205, 20);
            this.textBoxLJHF.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "恢复路径:";
            // 
            // btnRestore
            // 
            this.btnRestore.Location = new System.Drawing.Point(234, 52);
            this.btnRestore.Name = "btnRestore";
            this.btnRestore.Size = new System.Drawing.Size(75, 23);
            this.btnRestore.TabIndex = 1;
            this.btnRestore.Text = "恢复";
            this.btnRestore.UseVisualStyleBackColor = true;
            this.btnRestore.Click += new System.EventHandler(this.btnRestore_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnBKPath);
            this.groupBox2.Controls.Add(this.textBoxLJBK);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.btnBackUp);
            this.groupBox2.Location = new System.Drawing.Point(13, 133);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(315, 84);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "数据库备份";
            // 
            // btnBKPath
            // 
            this.btnBKPath.Location = new System.Drawing.Point(281, 20);
            this.btnBKPath.Name = "btnBKPath";
            this.btnBKPath.Size = new System.Drawing.Size(28, 23);
            this.btnBKPath.TabIndex = 7;
            this.btnBKPath.Text = "...";
            this.btnBKPath.UseVisualStyleBackColor = true;
            this.btnBKPath.Click += new System.EventHandler(this.btnBKPath_Click);
            // 
            // textBoxLJBK
            // 
            this.textBoxLJBK.Location = new System.Drawing.Point(70, 22);
            this.textBoxLJBK.Name = "textBoxLJBK";
            this.textBoxLJBK.Size = new System.Drawing.Size(205, 20);
            this.textBoxLJBK.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "备份路径:";
            // 
            // btnBackUp
            // 
            this.btnBackUp.Location = new System.Drawing.Point(234, 48);
            this.btnBackUp.Name = "btnBackUp";
            this.btnBackUp.Size = new System.Drawing.Size(75, 25);
            this.btnBackUp.TabIndex = 0;
            this.btnBackUp.Text = "备份";
            this.btnBackUp.UseVisualStyleBackColor = true;
            this.btnBackUp.Click += new System.EventHandler(this.btnBackUp_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 347);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(340, 22);
            this.statusStrip1.TabIndex = 8;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(320, 16);
            // 
            // saveFileDialogData
            // 
            this.saveFileDialogData.Filter = "所有(*.*)|*.*";
            this.saveFileDialogData.RestoreDirectory = true;
            // 
            // openFileDialogData
            // 
            this.openFileDialogData.Filter = "所有(*.*)|*.*";
            this.openFileDialogData.RestoreDirectory = true;
            // 
            // formDataMange
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(340, 369);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formDataMange";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "数据管理";
            this.Load += new System.EventHandler(this.formDataMange_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnPeopleClear;
        private System.Windows.Forms.Button btnConClear;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBoxLJHF;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRestore;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnBKPath;
        private System.Windows.Forms.TextBox textBoxLJBK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBackUp;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.SaveFileDialog saveFileDialogData;
        private System.Windows.Forms.OpenFileDialog openFileDialogData;
    }
}