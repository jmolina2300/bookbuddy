namespace BookBuddy
{
    partial class frm_patternmatch
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ed_srcColBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ed_srcKeyword = new System.Windows.Forms.TextBox();
            this.ed_destColBox1 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.ed_destColBox2 = new System.Windows.Forms.TextBox();
            this.ed_destColBox3 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.ed_destColBox4 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ed_srcColBox1);
            this.groupBox1.Controls.Add(this.ed_srcKeyword);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(189, 155);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Source";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBox4);
            this.groupBox2.Controls.Add(this.ed_destColBox1);
            this.groupBox2.Controls.Add(this.ed_destColBox4);
            this.groupBox2.Controls.Add(this.textBox10);
            this.groupBox2.Controls.Add(this.ed_destColBox3);
            this.groupBox2.Controls.Add(this.textBox8);
            this.groupBox2.Controls.Add(this.ed_destColBox2);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.textBox5);
            this.groupBox2.Location = new System.Drawing.Point(207, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(273, 155);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Destination(s)";
            // 
            // ed_srcColBox1
            // 
            this.ed_srcColBox1.Location = new System.Drawing.Point(12, 43);
            this.ed_srcColBox1.MaxLength = 5;
            this.ed_srcColBox1.Name = "ed_srcColBox1";
            this.ed_srcColBox1.Size = new System.Drawing.Size(31, 20);
            this.ed_srcColBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Column letter";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Keyword(s) to look for";
            // 
            // ed_srcKeyword
            // 
            this.ed_srcKeyword.Location = new System.Drawing.Point(12, 96);
            this.ed_srcKeyword.Name = "ed_srcKeyword";
            this.ed_srcKeyword.Size = new System.Drawing.Size(171, 20);
            this.ed_srcKeyword.TabIndex = 1;
            // 
            // ed_destColBox1
            // 
            this.ed_destColBox1.Location = new System.Drawing.Point(25, 43);
            this.ed_destColBox1.MaxLength = 5;
            this.ed_destColBox1.Name = "ed_destColBox1";
            this.ed_destColBox1.Size = new System.Drawing.Size(34, 20);
            this.ed_destColBox1.TabIndex = 2;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(88, 43);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(171, 20);
            this.textBox4.TabIndex = 3;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(89, 69);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(170, 20);
            this.textBox5.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(317, 173);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 32);
            this.button1.TabIndex = 2;
            this.button1.Text = "Go";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Column letter";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(85, 27);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Desired text";
            // 
            // ed_destColBox2
            // 
            this.ed_destColBox2.Location = new System.Drawing.Point(25, 69);
            this.ed_destColBox2.MaxLength = 5;
            this.ed_destColBox2.Name = "ed_destColBox2";
            this.ed_destColBox2.Size = new System.Drawing.Size(34, 20);
            this.ed_destColBox2.TabIndex = 4;
            // 
            // ed_destColBox3
            // 
            this.ed_destColBox3.Location = new System.Drawing.Point(25, 95);
            this.ed_destColBox3.MaxLength = 5;
            this.ed_destColBox3.Name = "ed_destColBox3";
            this.ed_destColBox3.Size = new System.Drawing.Size(34, 20);
            this.ed_destColBox3.TabIndex = 6;
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(89, 95);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(170, 20);
            this.textBox8.TabIndex = 7;
            // 
            // ed_destColBox4
            // 
            this.ed_destColBox4.Location = new System.Drawing.Point(25, 121);
            this.ed_destColBox4.MaxLength = 5;
            this.ed_destColBox4.Name = "ed_destColBox4";
            this.ed_destColBox4.Size = new System.Drawing.Size(34, 20);
            this.ed_destColBox4.TabIndex = 8;
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(89, 121);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(170, 20);
            this.textBox10.TabIndex = 9;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(398, 173);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 32);
            this.button2.TabIndex = 3;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // frm_patternmatch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(488, 212);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "frm_patternmatch";
            this.Text = "Keyword Matcher";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox ed_srcKeyword;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox ed_srcColBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox ed_destColBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox ed_destColBox2;
        private System.Windows.Forms.TextBox ed_destColBox4;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.TextBox ed_destColBox3;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.Button button2;
    }
}