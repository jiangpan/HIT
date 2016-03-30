namespace FG.PDMReader.Oracle
{
    partial class frmFGMDM
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
            this.dgvTabs = new System.Windows.Forms.DataGridView();
            this.btnGenScriptAll = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dgvTabCols = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTabs)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTabCols)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvTabs
            // 
            this.dgvTabs.AllowUserToAddRows = false;
            this.dgvTabs.AllowUserToDeleteRows = false;
            this.dgvTabs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTabs.Location = new System.Drawing.Point(3, 3);
            this.dgvTabs.MultiSelect = false;
            this.dgvTabs.Name = "dgvTabs";
            this.dgvTabs.ReadOnly = true;
            this.dgvTabs.RowTemplate.Height = 23;
            this.dgvTabs.Size = new System.Drawing.Size(258, 383);
            this.dgvTabs.TabIndex = 0;
            // 
            // btnGenScriptAll
            // 
            this.btnGenScriptAll.Location = new System.Drawing.Point(541, 448);
            this.btnGenScriptAll.Name = "btnGenScriptAll";
            this.btnGenScriptAll.Size = new System.Drawing.Size(75, 23);
            this.btnGenScriptAll.TabIndex = 1;
            this.btnGenScriptAll.Text = "button1";
            this.btnGenScriptAll.UseVisualStyleBackColor = true;
            this.btnGenScriptAll.Click += new System.EventHandler(this.btnGenScriptAll_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 453);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            // 
            // dgvTabCols
            // 
            this.dgvTabCols.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTabCols.Location = new System.Drawing.Point(267, 3);
            this.dgvTabCols.Name = "dgvTabCols";
            this.dgvTabCols.RowTemplate.Height = 23;
            this.dgvTabCols.Size = new System.Drawing.Size(462, 383);
            this.dgvTabCols.TabIndex = 3;
            // 
            // frmFGMDM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 510);
            this.Controls.Add(this.dgvTabCols);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnGenScriptAll);
            this.Controls.Add(this.dgvTabs);
            this.Name = "frmFGMDM";
            this.Text = "frmFGMDM";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmFGMDM_FormClosed);
            this.Load += new System.EventHandler(this.frmFGMDM_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTabs)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTabCols)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvTabs;
        private System.Windows.Forms.Button btnGenScriptAll;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgvTabCols;
    }
}