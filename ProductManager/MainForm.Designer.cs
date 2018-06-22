namespace ProductManager
{
    partial class MainForm
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
            this.loadExcelBtn = new System.Windows.Forms.Button();
            this.tipsText = new System.Windows.Forms.TextBox();
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // loadExcelBtn
            // 
            this.loadExcelBtn.Location = new System.Drawing.Point(24, 22);
            this.loadExcelBtn.Name = "loadExcelBtn";
            this.loadExcelBtn.Size = new System.Drawing.Size(473, 23);
            this.loadExcelBtn.TabIndex = 1;
            this.loadExcelBtn.Text = "LoadExcelFile";
            this.loadExcelBtn.UseVisualStyleBackColor = true;
            this.loadExcelBtn.Click += new System.EventHandler(this.loadExcelFileBtn_Click);
            // 
            // tipsText
            // 
            this.tipsText.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tipsText.Location = new System.Drawing.Point(12, 72);
            this.tipsText.Multiline = true;
            this.tipsText.Name = "tipsText";
            this.tipsText.Size = new System.Drawing.Size(879, 80);
            this.tipsText.TabIndex = 5;
            // 
            // webBrowser
            // 
            this.webBrowser.Location = new System.Drawing.Point(2, 168);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser.Size = new System.Drawing.Size(898, 340);
            this.webBrowser.TabIndex = 6;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(903, 513);
            this.Controls.Add(this.webBrowser);
            this.Controls.Add(this.tipsText);
            this.Controls.Add(this.loadExcelBtn);
            this.Name = "MainForm";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button loadExcelBtn;
        private System.Windows.Forms.TextBox tipsText;
        private System.Windows.Forms.WebBrowser webBrowser;
    }
}

