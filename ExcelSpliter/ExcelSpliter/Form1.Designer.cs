namespace ExcelSpliter
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dgvExcelList = new System.Windows.Forms.DataGridView();
            this.colFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnDone = new System.Windows.Forms.Button();
            this.txtFileCount = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRemove = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdSplitSheet = new System.Windows.Forms.RadioButton();
            this.rdSplitFile = new System.Windows.Forms.RadioButton();
            this.rdMergeToManySheet = new System.Windows.Forms.RadioButton();
            this.rdMergeToSingleSheet = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.lblProgress = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcelList)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvExcelList
            // 
            this.dgvExcelList.AllowUserToAddRows = false;
            this.dgvExcelList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvExcelList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colFileName});
            this.dgvExcelList.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvExcelList.Location = new System.Drawing.Point(0, 0);
            this.dgvExcelList.Name = "dgvExcelList";
            this.dgvExcelList.RowTemplate.Height = 23;
            this.dgvExcelList.Size = new System.Drawing.Size(570, 177);
            this.dgvExcelList.TabIndex = 0;
            // 
            // colFileName
            // 
            this.colFileName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colFileName.HeaderText = "文件名";
            this.colFileName.Name = "colFileName";
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(17, 183);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(79, 37);
            this.btnOpenFile.TabIndex = 1;
            this.btnOpenFile.Text = "打开文件";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // btnDone
            // 
            this.btnDone.Location = new System.Drawing.Point(443, 183);
            this.btnDone.Name = "btnDone";
            this.btnDone.Size = new System.Drawing.Size(79, 37);
            this.btnDone.TabIndex = 2;
            this.btnDone.Text = "执行";
            this.btnDone.UseVisualStyleBackColor = true;
            this.btnDone.Click += new System.EventHandler(this.btnDone_Click);
            // 
            // txtFileCount
            // 
            this.txtFileCount.Location = new System.Drawing.Point(319, 192);
            this.txtFileCount.Name = "txtFileCount";
            this.txtFileCount.Size = new System.Drawing.Size(118, 21);
            this.txtFileCount.TabIndex = 7;
            this.txtFileCount.Text = "2";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(206, 195);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "文件/工作区间个数:";
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(102, 183);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(79, 37);
            this.btnRemove.TabIndex = 9;
            this.btnRemove.Text = "移除文件";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdSplitSheet);
            this.groupBox1.Controls.Add(this.rdSplitFile);
            this.groupBox1.Controls.Add(this.rdMergeToManySheet);
            this.groupBox1.Controls.Add(this.rdMergeToSingleSheet);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Location = new System.Drawing.Point(0, 268);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(570, 97);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            // 
            // rdSplitSheet
            // 
            this.rdSplitSheet.AutoSize = true;
            this.rdSplitSheet.Location = new System.Drawing.Point(268, 59);
            this.rdSplitSheet.Name = "rdSplitSheet";
            this.rdSplitSheet.Size = new System.Drawing.Size(179, 16);
            this.rdSplitSheet.TabIndex = 10;
            this.rdSplitSheet.Text = "一个文件拆分成多个工作区间";
            this.rdSplitSheet.UseVisualStyleBackColor = true;
            // 
            // rdSplitFile
            // 
            this.rdSplitFile.AutoSize = true;
            this.rdSplitFile.Checked = true;
            this.rdSplitFile.Location = new System.Drawing.Point(268, 29);
            this.rdSplitFile.Name = "rdSplitFile";
            this.rdSplitFile.Size = new System.Drawing.Size(155, 16);
            this.rdSplitFile.TabIndex = 9;
            this.rdSplitFile.TabStop = true;
            this.rdSplitFile.Text = "一个文件拆分成多个文件";
            this.rdSplitFile.UseVisualStyleBackColor = true;
            // 
            // rdMergeToManySheet
            // 
            this.rdMergeToManySheet.AutoSize = true;
            this.rdMergeToManySheet.Location = new System.Drawing.Point(17, 59);
            this.rdMergeToManySheet.Name = "rdMergeToManySheet";
            this.rdMergeToManySheet.Size = new System.Drawing.Size(167, 16);
            this.rdMergeToManySheet.TabIndex = 8;
            this.rdMergeToManySheet.Text = "合并多文件成多个工作区间";
            this.rdMergeToManySheet.UseVisualStyleBackColor = true;
            // 
            // rdMergeToSingleSheet
            // 
            this.rdMergeToSingleSheet.AutoSize = true;
            this.rdMergeToSingleSheet.Location = new System.Drawing.Point(17, 29);
            this.rdMergeToSingleSheet.Name = "rdMergeToSingleSheet";
            this.rdMergeToSingleSheet.Size = new System.Drawing.Size(167, 16);
            this.rdMergeToSingleSheet.TabIndex = 7;
            this.rdMergeToSingleSheet.Text = "合并多文件成一个工作区间";
            this.rdMergeToSingleSheet.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 244);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 12);
            this.label2.TabIndex = 11;
            this.label2.Text = "进度:";
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(59, 245);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(23, 12);
            this.lblProgress.TabIndex = 12;
            this.lblProgress.Text = "0/0";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 365);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFileCount);
            this.Controls.Add(this.btnDone);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.dgvExcelList);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel合并分割神器";
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcelList)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvExcelList;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFileName;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnDone;
        private System.Windows.Forms.TextBox txtFileCount;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdSplitSheet;
        private System.Windows.Forms.RadioButton rdSplitFile;
        private System.Windows.Forms.RadioButton rdMergeToManySheet;
        private System.Windows.Forms.RadioButton rdMergeToSingleSheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblProgress;
    }
}

