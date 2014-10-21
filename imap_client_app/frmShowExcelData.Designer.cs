namespace imap_client_app
{
    partial class frmShowExcelData
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
            this.dg = new System.Windows.Forms.DataGridView();
            this.ad_data1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ad_data2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ad_data3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dg)).BeginInit();
            this.SuspendLayout();
            // 
            // dg
            // 
            this.dg.AllowUserToAddRows = false;
            this.dg.AllowUserToOrderColumns = true;
            this.dg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ad_data1,
            this.ad_data2,
            this.ad_data3});
            this.dg.Location = new System.Drawing.Point(12, 38);
            this.dg.Name = "dg";
            this.dg.Size = new System.Drawing.Size(584, 276);
            this.dg.TabIndex = 0;
            // 
            // ad_data1
            // 
            this.ad_data1.DataPropertyName = "ad_data1";
            this.ad_data1.HeaderText = "Data1";
            this.ad_data1.Name = "ad_data1";
            this.ad_data1.ReadOnly = true;
            // 
            // ad_data2
            // 
            this.ad_data2.DataPropertyName = "ad_data2";
            this.ad_data2.HeaderText = "Data2";
            this.ad_data2.Name = "ad_data2";
            this.ad_data2.ReadOnly = true;
            // 
            // ad_data3
            // 
            this.ad_data3.DataPropertyName = "ad_data3";
            this.ad_data3.HeaderText = "Data3";
            this.ad_data3.Name = "ad_data3";
            this.ad_data3.ReadOnly = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Search";
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(59, 12);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(254, 20);
            this.txtSearch.TabIndex = 2;
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(521, 9);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmShowExcelData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(609, 326);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.txtSearch);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dg);
            this.Name = "frmShowExcelData";
            this.Text = "frmShowExcelData";
            this.Load += new System.EventHandler(this.frmShowExcelData_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.DataGridViewTextBoxColumn ad_data1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ad_data2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ad_data3;
        private System.Windows.Forms.Button btnCancel;
    }
}