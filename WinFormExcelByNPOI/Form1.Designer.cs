namespace WinFormExcelByNPOI
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
            this.tbFilePath = new System.Windows.Forms.TextBox();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnChoseFile = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbFilePath
            // 
            this.tbFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbFilePath.Location = new System.Drawing.Point(12, 12);
            this.tbFilePath.Name = "tbFilePath";
            this.tbFilePath.Size = new System.Drawing.Size(382, 28);
            this.tbFilePath.TabIndex = 5;
            // 
            // btnImport
            // 
            this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImport.Location = new System.Drawing.Point(142, 53);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(123, 44);
            this.btnImport.TabIndex = 4;
            this.btnImport.Text = "导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnChoseFile
            // 
            this.btnChoseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnChoseFile.Location = new System.Drawing.Point(12, 53);
            this.btnChoseFile.Name = "btnChoseFile";
            this.btnChoseFile.Size = new System.Drawing.Size(124, 44);
            this.btnChoseFile.TabIndex = 3;
            this.btnChoseFile.Text = "选择文件";
            this.btnChoseFile.UseVisualStyleBackColor = true;
            this.btnChoseFile.Click += new System.EventHandler(this.btnChoseFile_Click);
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.Location = new System.Drawing.Point(271, 53);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(123, 44);
            this.btnExport.TabIndex = 6;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 109);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.tbFilePath);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnChoseFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbFilePath;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnChoseFile;
        private System.Windows.Forms.Button btnExport;
    }
}

