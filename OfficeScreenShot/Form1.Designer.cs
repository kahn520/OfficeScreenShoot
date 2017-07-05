namespace OfficeScreenShot
{
    partial class Form1
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ColFile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPage = new System.Windows.Forms.TextBox();
            this.radioPpt = new System.Windows.Forms.RadioButton();
            this.radioDoc = new System.Windows.Forms.RadioButton();
            this.radioXls = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(3, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(56, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "选择";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtFolder
            // 
            this.txtFolder.Location = new System.Drawing.Point(65, 5);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(229, 21);
            this.txtFolder.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColFile,
            this.ColStatus});
            this.dataGridView1.Location = new System.Drawing.Point(3, 61);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(476, 529);
            this.dataGridView1.TabIndex = 2;
            // 
            // ColFile
            // 
            this.ColFile.DataPropertyName = "file";
            this.ColFile.FillWeight = 159.3909F;
            this.ColFile.HeaderText = "文件名";
            this.ColFile.Name = "ColFile";
            this.ColFile.ReadOnly = true;
            // 
            // ColStatus
            // 
            this.ColStatus.DataPropertyName = "status";
            this.ColStatus.FillWeight = 40.60914F;
            this.ColStatus.HeaderText = "状态";
            this.ColStatus.Name = "ColStatus";
            this.ColStatus.ReadOnly = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(3, 32);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(56, 23);
            this.button2.TabIndex = 0;
            this.button2.Text = "开始";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(334, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "页数";
            // 
            // txtPage
            // 
            this.txtPage.Location = new System.Drawing.Point(369, 5);
            this.txtPage.Name = "txtPage";
            this.txtPage.Size = new System.Drawing.Size(43, 21);
            this.txtPage.TabIndex = 4;
            this.txtPage.Text = "99999";
            // 
            // radioPpt
            // 
            this.radioPpt.AutoSize = true;
            this.radioPpt.Location = new System.Drawing.Point(128, 35);
            this.radioPpt.Name = "radioPpt";
            this.radioPpt.Size = new System.Drawing.Size(47, 16);
            this.radioPpt.TabIndex = 5;
            this.radioPpt.Text = ".ppt";
            this.radioPpt.UseVisualStyleBackColor = true;
            // 
            // radioDoc
            // 
            this.radioDoc.AutoSize = true;
            this.radioDoc.Checked = true;
            this.radioDoc.Location = new System.Drawing.Point(195, 35);
            this.radioDoc.Name = "radioDoc";
            this.radioDoc.Size = new System.Drawing.Size(47, 16);
            this.radioDoc.TabIndex = 5;
            this.radioDoc.TabStop = true;
            this.radioDoc.Text = ".doc";
            this.radioDoc.UseVisualStyleBackColor = true;
            // 
            // radioXls
            // 
            this.radioXls.AutoSize = true;
            this.radioXls.Location = new System.Drawing.Point(263, 35);
            this.radioXls.Name = "radioXls";
            this.radioXls.Size = new System.Drawing.Size(47, 16);
            this.radioXls.TabIndex = 5;
            this.radioXls.Text = ".xls";
            this.radioXls.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 590);
            this.Controls.Add(this.radioXls);
            this.Controls.Add(this.radioDoc);
            this.Controls.Add(this.radioPpt);
            this.Controls.Add(this.txtPage);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "截图";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPage;
        private System.Windows.Forms.RadioButton radioPpt;
        private System.Windows.Forms.RadioButton radioDoc;
        private System.Windows.Forms.RadioButton radioXls;
    }
}

