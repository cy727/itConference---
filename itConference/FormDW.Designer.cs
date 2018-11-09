namespace itConference
{
    partial class FormDW
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormDW));
            this.btnADD = new System.Windows.Forms.Button();
            this.btnEDIT = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnDEL = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.treeViewComm = new System.Windows.Forms.TreeView();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxDWMC = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxSJDW = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnADD
            // 
            this.btnADD.Location = new System.Drawing.Point(15, 476);
            this.btnADD.Name = "btnADD";
            this.btnADD.Size = new System.Drawing.Size(77, 25);
            this.btnADD.TabIndex = 15;
            this.btnADD.Text = "增加";
            this.btnADD.UseVisualStyleBackColor = true;
            this.btnADD.Click += new System.EventHandler(this.btnADD_Click);
            // 
            // btnEDIT
            // 
            this.btnEDIT.Location = new System.Drawing.Point(164, 476);
            this.btnEDIT.Name = "btnEDIT";
            this.btnEDIT.Size = new System.Drawing.Size(63, 25);
            this.btnEDIT.TabIndex = 14;
            this.btnEDIT.Text = "修改";
            this.btnEDIT.UseVisualStyleBackColor = true;
            this.btnEDIT.Click += new System.EventHandler(this.btnEDIT_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(263, 476);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnDEL
            // 
            this.btnDEL.Location = new System.Drawing.Point(98, 476);
            this.btnDEL.Name = "btnDEL";
            this.btnDEL.Size = new System.Drawing.Size(60, 25);
            this.btnDEL.TabIndex = 12;
            this.btnDEL.Text = "删除";
            this.btnDEL.UseVisualStyleBackColor = true;
            this.btnDEL.Click += new System.EventHandler(this.btnDEL_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.treeViewComm);
            this.groupBox1.Location = new System.Drawing.Point(12, 95);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(329, 375);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "单位";
            // 
            // treeViewComm
            // 
            this.treeViewComm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewComm.Location = new System.Drawing.Point(3, 16);
            this.treeViewComm.Name = "treeViewComm";
            this.treeViewComm.Size = new System.Drawing.Size(323, 356);
            this.treeViewComm.TabIndex = 0;
            this.treeViewComm.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewComm_AfterSelect);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "单位名称:";
            // 
            // textBoxDWMC
            // 
            this.textBoxDWMC.Location = new System.Drawing.Point(87, 19);
            this.textBoxDWMC.Name = "textBoxDWMC";
            this.textBoxDWMC.Size = new System.Drawing.Size(228, 20);
            this.textBoxDWMC.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "上级单位:";
            // 
            // textBoxSJDW
            // 
            this.textBoxSJDW.Location = new System.Drawing.Point(87, 47);
            this.textBoxSJDW.Name = "textBoxSJDW";
            this.textBoxSJDW.ReadOnly = true;
            this.textBoxSJDW.Size = new System.Drawing.Size(228, 20);
            this.textBoxSJDW.TabIndex = 3;
            this.textBoxSJDW.DoubleClick += new System.EventHandler(this.textBoxSJDW_DoubleClick);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBoxSJDW);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.textBoxDWMC);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(12, 9);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(329, 80);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "单位明细";
            // 
            // FormDW
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(354, 513);
            this.Controls.Add(this.btnADD);
            this.Controls.Add(this.btnEDIT);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDEL);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormDW";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "单位关系维护";
            this.Load += new System.EventHandler(this.FormDW_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnADD;
        private System.Windows.Forms.Button btnEDIT;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnDEL;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TreeView treeViewComm;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxDWMC;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxSJDW;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}