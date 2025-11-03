namespace BookBuddy
{
    partial class frmDescriptionAutofill
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
            this.label5 = new System.Windows.Forms.Label();
            this.txtSourceColumn = new System.Windows.Forms.TextBox();
            this.txtDestinationColumn = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnCancel2 = new System.Windows.Forms.Button();
            this.btnOKMIMO = new System.Windows.Forms.Button();
            this.btnImportCSV = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Source = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkUseOldMatchingAlgorithm = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.Source.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 23);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(42, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Column";
            // 
            // txtSourceColumn
            // 
            this.txtSourceColumn.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtSourceColumn.Location = new System.Drawing.Point(58, 20);
            this.txtSourceColumn.Name = "txtSourceColumn";
            this.txtSourceColumn.Size = new System.Drawing.Size(26, 20);
            this.txtSourceColumn.TabIndex = 1;
            this.txtSourceColumn.TextChanged += new System.EventHandler(this.txtSourceColumn_TextChanged);
            // 
            // txtDestinationColumn
            // 
            this.txtDestinationColumn.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtDestinationColumn.Location = new System.Drawing.Point(54, 20);
            this.txtDestinationColumn.Name = "txtDestinationColumn";
            this.txtDestinationColumn.Size = new System.Drawing.Size(26, 20);
            this.txtDestinationColumn.TabIndex = 3;
            this.txtDestinationColumn.TextChanged += new System.EventHandler(this.txtDestinationColumn_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 23);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 9;
            this.label6.Text = "Column";
            // 
            // btnCancel2
            // 
            this.btnCancel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel2.Location = new System.Drawing.Point(381, 411);
            this.btnCancel2.Name = "btnCancel2";
            this.btnCancel2.Size = new System.Drawing.Size(75, 23);
            this.btnCancel2.TabIndex = 3;
            this.btnCancel2.Text = "Cancel";
            this.btnCancel2.UseVisualStyleBackColor = true;
            this.btnCancel2.Click += new System.EventHandler(this.btnCancel2_Click);
            // 
            // btnOKMIMO
            // 
            this.btnOKMIMO.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOKMIMO.Location = new System.Drawing.Point(300, 411);
            this.btnOKMIMO.Name = "btnOKMIMO";
            this.btnOKMIMO.Size = new System.Drawing.Size(75, 23);
            this.btnOKMIMO.TabIndex = 2;
            this.btnOKMIMO.Text = "OK";
            this.btnOKMIMO.UseVisualStyleBackColor = true;
            this.btnOKMIMO.Click += new System.EventHandler(this.btnOKMIMO_Click);
            // 
            // btnImportCSV
            // 
            this.btnImportCSV.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnImportCSV.Location = new System.Drawing.Point(14, 411);
            this.btnImportCSV.Name = "btnImportCSV";
            this.btnImportCSV.Size = new System.Drawing.Size(96, 23);
            this.btnImportCSV.TabIndex = 1;
            this.btnImportCSV.Text = "Import CSV";
            this.btnImportCSV.UseVisualStyleBackColor = true;
            this.btnImportCSV.Click += new System.EventHandler(this.btnImportCSV_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2});
            this.dataGridView1.Location = new System.Drawing.Point(14, 116);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(442, 279);
            this.dataGridView1.TabIndex = 0;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Keyword";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Autofill Description";
            this.Column2.Name = "Column2";
            // 
            // Source
            // 
            this.Source.Controls.Add(this.label5);
            this.Source.Controls.Add(this.txtSourceColumn);
            this.Source.Location = new System.Drawing.Point(14, 19);
            this.Source.Name = "Source";
            this.Source.Size = new System.Drawing.Size(122, 50);
            this.Source.TabIndex = 10;
            this.Source.TabStop = false;
            this.Source.Text = "Source";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.txtDestinationColumn);
            this.groupBox3.Location = new System.Drawing.Point(168, 19);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(113, 50);
            this.groupBox3.TabIndex = 11;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Destination";
            // 
            // chkUseOldMatchingAlgorithm
            // 
            this.chkUseOldMatchingAlgorithm.AutoSize = true;
            this.chkUseOldMatchingAlgorithm.Location = new System.Drawing.Point(14, 93);
            this.chkUseOldMatchingAlgorithm.Name = "chkUseOldMatchingAlgorithm";
            this.chkUseOldMatchingAlgorithm.Size = new System.Drawing.Size(153, 17);
            this.chkUseOldMatchingAlgorithm.TabIndex = 12;
            this.chkUseOldMatchingAlgorithm.Text = "Use old matching algorithm";
            this.chkUseOldMatchingAlgorithm.UseVisualStyleBackColor = true;
            // 
            // frmDescriptionAutofill
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 446);
            this.Controls.Add(this.chkUseOldMatchingAlgorithm);
            this.Controls.Add(this.btnCancel2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.btnOKMIMO);
            this.Controls.Add(this.Source);
            this.Controls.Add(this.btnImportCSV);
            this.Controls.Add(this.dataGridView1);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(480, 480);
            this.Name = "frmDescriptionAutofill";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Multiple Description Autofill";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.Source.ResumeLayout(false);
            this.Source.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSourceColumn;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtDestinationColumn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnImportCSV;
        private System.Windows.Forms.GroupBox Source;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnOKMIMO;
        private System.Windows.Forms.Button btnCancel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.CheckBox chkUseOldMatchingAlgorithm;
    }
}