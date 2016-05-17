namespace FG.PDMReader
{
    partial class frmMain
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
            this.btnGenScript = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnBrowser = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.mnsMain = new System.Windows.Forms.MenuStrip();
            this.pDM操作ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiOracle = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiOracleFGMDM = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiMSSQL = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiMSSQLFGMDMDict = new System.Windows.Forms.ToolStripMenuItem();
            this.mnsMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenScript
            // 
            this.btnGenScript.Location = new System.Drawing.Point(624, 115);
            this.btnGenScript.Name = "btnGenScript";
            this.btnGenScript.Size = new System.Drawing.Size(75, 23);
            this.btnGenScript.TabIndex = 1;
            this.btnGenScript.Text = "生成脚本";
            this.btnGenScript.UseVisualStyleBackColor = true;
            this.btnGenScript.Click += new System.EventHandler(this.btnGenScript_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(4, 115);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(608, 487);
            this.richTextBox1.TabIndex = 2;
            this.richTextBox1.Text = "";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(8, 57);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(608, 21);
            this.textBox1.TabIndex = 3;
            // 
            // btnBrowser
            // 
            this.btnBrowser.Location = new System.Drawing.Point(624, 55);
            this.btnBrowser.Name = "btnBrowser";
            this.btnBrowser.Size = new System.Drawing.Size(75, 23);
            this.btnBrowser.TabIndex = 4;
            this.btnBrowser.Text = "浏览";
            this.btnBrowser.UseVisualStyleBackColor = true;
            this.btnBrowser.Click += new System.EventHandler(this.btnBrowser_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(624, 173);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 5;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnGenExcel_Click);
            // 
            // mnsMain
            // 
            this.mnsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pDM操作ToolStripMenuItem,
            this.tsmiOracle,
            this.tsmiMSSQL});
            this.mnsMain.Location = new System.Drawing.Point(0, 0);
            this.mnsMain.Name = "mnsMain";
            this.mnsMain.Size = new System.Drawing.Size(711, 25);
            this.mnsMain.TabIndex = 6;
            this.mnsMain.Text = "menuStrip1";
            // 
            // pDM操作ToolStripMenuItem
            // 
            this.pDM操作ToolStripMenuItem.Name = "pDM操作ToolStripMenuItem";
            this.pDM操作ToolStripMenuItem.Size = new System.Drawing.Size(72, 21);
            this.pDM操作ToolStripMenuItem.Text = "PDM操作";
            // 
            // tsmiOracle
            // 
            this.tsmiOracle.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiOracleFGMDM});
            this.tsmiOracle.Name = "tsmiOracle";
            this.tsmiOracle.Size = new System.Drawing.Size(94, 21);
            this.tsmiOracle.Text = "Oracle数据库";
            // 
            // tsmiOracleFGMDM
            // 
            this.tsmiOracleFGMDM.Name = "tsmiOracleFGMDM";
            this.tsmiOracleFGMDM.Size = new System.Drawing.Size(226, 22);
            this.tsmiOracleFGMDM.Text = "FGMDM库生成MSSQL脚本";
            this.tsmiOracleFGMDM.Click += new System.EventHandler(this.tsmiOracleFGMDM_Click);
            // 
            // tsmiMSSQL
            // 
            this.tsmiMSSQL.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiMSSQLFGMDMDict});
            this.tsmiMSSQL.Name = "tsmiMSSQL";
            this.tsmiMSSQL.Size = new System.Drawing.Size(98, 21);
            this.tsmiMSSQL.Text = "MSSQL数据库";
            // 
            // tsmiMSSQLFGMDMDict
            // 
            this.tsmiMSSQLFGMDMDict.Name = "tsmiMSSQLFGMDMDict";
            this.tsmiMSSQLFGMDMDict.Size = new System.Drawing.Size(172, 22);
            this.tsmiMSSQLFGMDMDict.Text = "FGMDM数据字典";
            this.tsmiMSSQLFGMDMDict.Click += new System.EventHandler(this.tsmiMSSQLFGMDMDict_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(711, 615);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnBrowser);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.btnGenScript);
            this.Controls.Add(this.mnsMain);
            this.MainMenuStrip = this.mnsMain;
            this.Name = "frmMain";
            this.Text = "Form1";
            this.mnsMain.ResumeLayout(false);
            this.mnsMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenScript;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnBrowser;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.MenuStrip mnsMain;
        private System.Windows.Forms.ToolStripMenuItem pDM操作ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiOracle;
        private System.Windows.Forms.ToolStripMenuItem tsmiOracleFGMDM;
        private System.Windows.Forms.ToolStripMenuItem tsmiMSSQL;
        private System.Windows.Forms.ToolStripMenuItem tsmiMSSQLFGMDMDict;
    }
}

