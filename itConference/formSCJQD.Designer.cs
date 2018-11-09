namespace itConference
{
    partial class formSCJQD
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formSCJQD));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBoxSJ = new System.Windows.Forms.CheckBox();
            this.checkBoxAll = new System.Windows.Forms.CheckBox();
            this.textBoxXM = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dataGridViewMX = new System.Windows.Forms.DataGridView();
            this.buttonPrn = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.sqlDA = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMX)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.checkBoxSJ);
            this.groupBox1.Controls.Add(this.checkBoxAll);
            this.groupBox1.Controls.Add(this.textBoxXM);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(713, 59);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "条件过滤";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(632, 25);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "过滤整理";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBoxSJ
            // 
            this.checkBoxSJ.AutoSize = true;
            this.checkBoxSJ.Checked = true;
            this.checkBoxSJ.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSJ.Location = new System.Drawing.Point(521, 28);
            this.checkBoxSJ.Name = "checkBoxSJ";
            this.checkBoxSJ.Size = new System.Drawing.Size(96, 16);
            this.checkBoxSJ.TabIndex = 4;
            this.checkBoxSJ.Text = "包含签到时间";
            this.checkBoxSJ.UseVisualStyleBackColor = true;
            // 
            // checkBoxAll
            // 
            this.checkBoxAll.AutoSize = true;
            this.checkBoxAll.Checked = true;
            this.checkBoxAll.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxAll.Location = new System.Drawing.Point(428, 28);
            this.checkBoxAll.Name = "checkBoxAll";
            this.checkBoxAll.Size = new System.Drawing.Size(72, 16);
            this.checkBoxAll.TabIndex = 2;
            this.checkBoxAll.Text = "所有人员";
            this.checkBoxAll.UseVisualStyleBackColor = true;
            // 
            // textBoxXM
            // 
            this.textBoxXM.Location = new System.Drawing.Point(63, 26);
            this.textBoxXM.Name = "textBoxXM";
            this.textBoxXM.Size = new System.Drawing.Size(343, 21);
            this.textBoxXM.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "姓名：";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dataGridViewMX);
            this.groupBox2.Location = new System.Drawing.Point(12, 77);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(713, 339);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "列表";
            // 
            // dataGridViewMX
            // 
            this.dataGridViewMX.AllowUserToAddRows = false;
            this.dataGridViewMX.AllowUserToDeleteRows = false;
            this.dataGridViewMX.AllowUserToOrderColumns = true;
            this.dataGridViewMX.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMX.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewMX.Location = new System.Drawing.Point(3, 17);
            this.dataGridViewMX.Name = "dataGridViewMX";
            this.dataGridViewMX.RowTemplate.Height = 23;
            this.dataGridViewMX.Size = new System.Drawing.Size(707, 319);
            this.dataGridViewMX.TabIndex = 0;
            // 
            // buttonPrn
            // 
            this.buttonPrn.Location = new System.Drawing.Point(503, 422);
            this.buttonPrn.Name = "buttonPrn";
            this.buttonPrn.Size = new System.Drawing.Size(135, 23);
            this.buttonPrn.TabIndex = 2;
            this.buttonPrn.Text = "打印";
            this.buttonPrn.UseVisualStyleBackColor = true;
            this.buttonPrn.Click += new System.EventHandler(this.buttonPrn_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(644, 422);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(75, 23);
            this.buttonClose.TabIndex = 3;
            this.buttonClose.Text = "关闭";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 461);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(737, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // sqlDA
            // 
            this.sqlDA.DeleteCommand = this.sqlDeleteCommand1;
            this.sqlDA.InsertCommand = this.sqlInsertCommand1;
            this.sqlDA.SelectCommand = this.sqlSelectCommand1;
            this.sqlDA.UpdateCommand = this.sqlUpdateCommand1;
            // 
            // sqlConn
            // 
            this.sqlConn.FireInfoMessageEventOnUserErrors = false;
            // 
            // formSCJQD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(737, 483);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonPrn);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formSCJQD";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "手持机签到信息";
            this.Load += new System.EventHandler(this.formSCJQD_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMX)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBoxSJ;
        private System.Windows.Forms.CheckBox checkBoxAll;
        private System.Windows.Forms.TextBox textBoxXM;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dataGridViewMX;
        private System.Windows.Forms.Button buttonPrn;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Data.SqlClient.SqlDataAdapter sqlDA;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Data.SqlClient.SqlConnection sqlConn;
    }
}