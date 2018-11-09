namespace itConference
{
    partial class formTrainManage
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formTrainManage));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dataGridViewTrain = new System.Windows.Forms.DataGridView();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSign = new System.Windows.Forms.Button();
            this.btnResign = new System.Windows.Forms.Button();
            this.btnOffsign = new System.Windows.Forms.Button();
            this.btnCount = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTrain)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dataGridViewTrain);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(765, 488);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "培训";
            // 
            // dataGridViewTrain
            // 
            this.dataGridViewTrain.AllowUserToAddRows = false;
            this.dataGridViewTrain.AllowUserToDeleteRows = false;
            this.dataGridViewTrain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewTrain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewTrain.Location = new System.Drawing.Point(3, 16);
            this.dataGridViewTrain.Name = "dataGridViewTrain";
            this.dataGridViewTrain.ReadOnly = true;
            this.dataGridViewTrain.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewTrain.Size = new System.Drawing.Size(759, 469);
            this.dataGridViewTrain.TabIndex = 0;
            this.dataGridViewTrain.DoubleClick += new System.EventHandler(this.dataGridViewTrain_DoubleClick);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnAdd.Location = new System.Drawing.Point(44, 510);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 26);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Text = "添加会议";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnEdit.Location = new System.Drawing.Point(125, 510);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 25);
            this.btnEdit.TabIndex = 4;
            this.btnEdit.Text = "修改会议";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnDel.Location = new System.Drawing.Point(206, 510);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 25);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "删除会议";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(675, 511);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnSign
            // 
            this.btnSign.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnSign.Location = new System.Drawing.Point(44, 511);
            this.btnSign.Name = "btnSign";
            this.btnSign.Size = new System.Drawing.Size(75, 25);
            this.btnSign.TabIndex = 9;
            this.btnSign.Text = "会议签到";
            this.btnSign.UseVisualStyleBackColor = true;
            this.btnSign.Click += new System.EventHandler(this.btnSign_Click);
            // 
            // btnResign
            // 
            this.btnResign.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnResign.Location = new System.Drawing.Point(125, 510);
            this.btnResign.Name = "btnResign";
            this.btnResign.Size = new System.Drawing.Size(75, 25);
            this.btnResign.TabIndex = 8;
            this.btnResign.Text = "忘卡补签";
            this.btnResign.UseVisualStyleBackColor = true;
            this.btnResign.Click += new System.EventHandler(this.btnResign_Click);
            // 
            // btnOffsign
            // 
            this.btnOffsign.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnOffsign.Location = new System.Drawing.Point(206, 511);
            this.btnOffsign.Name = "btnOffsign";
            this.btnOffsign.Size = new System.Drawing.Size(75, 25);
            this.btnOffsign.TabIndex = 7;
            this.btnOffsign.Text = "离线签到";
            this.btnOffsign.UseVisualStyleBackColor = true;
            this.btnOffsign.Click += new System.EventHandler(this.btnOffsign_Click);
            // 
            // btnCount
            // 
            this.btnCount.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnCount.Location = new System.Drawing.Point(44, 512);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(75, 23);
            this.btnCount.TabIndex = 10;
            this.btnCount.Text = "统计输出";
            this.btnCount.UseVisualStyleBackColor = true;
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // formTrainManage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(790, 545);
            this.Controls.Add(this.btnOffsign);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.btnResign);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnSign);
            this.Controls.Add(this.btnAdd);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "formTrainManage";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "培训定义";
            this.Load += new System.EventHandler(this.formTrainManage_Load);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTrain)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.DataGridView dataGridViewTrain;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSign;
        private System.Windows.Forms.Button btnResign;
        private System.Windows.Forms.Button btnOffsign;
        private System.Windows.Forms.Button btnCount;
    }
}