namespace itConference
{
    partial class formTrainResign
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formTrainResign));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelConferencename = new System.Windows.Forms.Label();
            this.dataGridViewreSign = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnreSign = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelresign = new System.Windows.Forms.ToolStripStatusLabel();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlComm = new System.Data.SqlClient.SqlCommand();
            this.sqlDA = new System.Data.SqlClient.SqlDataAdapter();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDownBQ = new System.Windows.Forms.NumericUpDown();
            this.btnBQ = new System.Windows.Forms.Button();
            this.btnQB = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewreSign)).BeginInit();
            this.panel2.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBQ)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridViewreSign, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.statusStrip1, 0, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(639, 489);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Azure;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.labelConferencename);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(633, 34);
            this.panel1.TabIndex = 0;
            // 
            // labelConferencename
            // 
            this.labelConferencename.AutoSize = true;
            this.labelConferencename.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelConferencename.Location = new System.Drawing.Point(9, 5);
            this.labelConferencename.Name = "labelConferencename";
            this.labelConferencename.Size = new System.Drawing.Size(75, 19);
            this.labelConferencename.TabIndex = 0;
            this.labelConferencename.Text = "label1";
            // 
            // dataGridViewreSign
            // 
            this.dataGridViewreSign.AllowUserToAddRows = false;
            this.dataGridViewreSign.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewreSign.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewreSign.Location = new System.Drawing.Point(3, 43);
            this.dataGridViewreSign.Name = "dataGridViewreSign";
            this.dataGridViewreSign.ReadOnly = true;
            this.dataGridViewreSign.RowTemplate.Height = 23;
            this.dataGridViewreSign.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewreSign.Size = new System.Drawing.Size(633, 373);
            this.dataGridViewreSign.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.btnQB);
            this.panel2.Controls.Add(this.btnBQ);
            this.panel2.Controls.Add(this.numericUpDownBQ);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.btnCancel);
            this.panel2.Controls.Add(this.btnreSign);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 422);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(633, 34);
            this.panel2.TabIndex = 2;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(542, 7);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnreSign
            // 
            this.btnreSign.Location = new System.Drawing.Point(461, 7);
            this.btnreSign.Name = "btnreSign";
            this.btnreSign.Size = new System.Drawing.Size(75, 23);
            this.btnreSign.TabIndex = 0;
            this.btnreSign.Text = "确定";
            this.btnreSign.UseVisualStyleBackColor = true;
            this.btnreSign.Click += new System.EventHandler(this.btnreSign_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelresign});
            this.statusStrip1.Location = new System.Drawing.Point(0, 467);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(639, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabelresign
            // 
            this.toolStripStatusLabelresign.Name = "toolStripStatusLabelresign";
            this.toolStripStatusLabelresign.Size = new System.Drawing.Size(131, 17);
            this.toolStripStatusLabelresign.Text = "toolStripStatusLabel1";
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "补签人数：";
            // 
            // numericUpDownBQ
            // 
            this.numericUpDownBQ.Location = new System.Drawing.Point(69, 10);
            this.numericUpDownBQ.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numericUpDownBQ.Name = "numericUpDownBQ";
            this.numericUpDownBQ.Size = new System.Drawing.Size(80, 20);
            this.numericUpDownBQ.TabIndex = 3;
            this.numericUpDownBQ.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnBQ
            // 
            this.btnBQ.Location = new System.Drawing.Point(155, 7);
            this.btnBQ.Name = "btnBQ";
            this.btnBQ.Size = new System.Drawing.Size(75, 23);
            this.btnBQ.TabIndex = 4;
            this.btnBQ.Text = "补签";
            this.btnBQ.UseVisualStyleBackColor = true;
            this.btnBQ.Click += new System.EventHandler(this.btnBQ_Click);
            // 
            // btnQB
            // 
            this.btnQB.Location = new System.Drawing.Point(236, 7);
            this.btnQB.Name = "btnQB";
            this.btnQB.Size = new System.Drawing.Size(75, 23);
            this.btnQB.TabIndex = 5;
            this.btnQB.Text = "全部补签";
            this.btnQB.UseVisualStyleBackColor = true;
            this.btnQB.Click += new System.EventHandler(this.btnQB_Click);
            // 
            // formTrainResign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(639, 489);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formTrainResign";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "忘卡补签";
            this.Load += new System.EventHandler(this.formTrainResign_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewreSign)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBQ)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelConferencename;
        private System.Windows.Forms.DataGridView dataGridViewreSign;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnreSign;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelresign;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlComm;
        private System.Data.SqlClient.SqlDataAdapter sqlDA;
        private System.Windows.Forms.Button btnQB;
        private System.Windows.Forms.Button btnBQ;
        private System.Windows.Forms.NumericUpDown numericUpDownBQ;
        private System.Windows.Forms.Label label1;
    }
}