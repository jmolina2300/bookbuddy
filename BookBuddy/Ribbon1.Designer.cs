namespace BookBuddy
{
    partial class Ribbon1
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem3 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem4 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.box4 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.label1 = new Microsoft.Office.Tools.Ribbon.RibbonLabel();
            this.box1 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.ed_colBox1 = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.ed_textBox1 = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.separator1 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.box3 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.label2 = new Microsoft.Office.Tools.Ribbon.RibbonLabel();
            this.box2 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.ed_colBox2 = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.ed_textBox2 = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.separator2 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.btn_go = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.group2 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.box5 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.label3 = new Microsoft.Office.Tools.Ribbon.RibbonLabel();
            this.box6 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.ed_signFlip_colBox = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.cb_pickSign = new Microsoft.Office.Tools.Ribbon.RibbonComboBox();
            this.separator3 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.btn_go_signFlip = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.group3 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.box7 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.label4 = new Microsoft.Office.Tools.Ribbon.RibbonLabel();
            this.box8 = new Microsoft.Office.Tools.Ribbon.RibbonBox();
            this.ed_cellCleanup_column = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.ed_cellCleanup_characters = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.separator4 = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
            this.btn_go_cellCleanup = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box4.SuspendLayout();
            this.box1.SuspendLayout();
            this.box3.SuspendLayout();
            this.box2.SuspendLayout();
            this.group2.SuspendLayout();
            this.box5.SuspendLayout();
            this.box6.SuspendLayout();
            this.group3.SuspendLayout();
            this.box7.SuspendLayout();
            this.box8.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "BookBuddy";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.box4);
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.box3);
            this.group1.Items.Add(this.box2);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.btn_go);
            this.group1.Label = "Description Auto Filler";
            this.group1.Name = "group1";
            // 
            // box4
            // 
            this.box4.Items.Add(this.label1);
            this.box4.Name = "box4";
            // 
            // label1
            // 
            this.label1.Label = "Source";
            this.label1.Name = "label1";
            // 
            // box1
            // 
            this.box1.Items.Add(this.ed_colBox1);
            this.box1.Items.Add(this.ed_textBox1);
            this.box1.Name = "box1";
            // 
            // ed_colBox1
            // 
            this.ed_colBox1.Label = "Column";
            this.ed_colBox1.Name = "ed_colBox1";
            this.ed_colBox1.SizeString = "aaa";
            this.ed_colBox1.Text = null;
            // 
            // ed_textBox1
            // 
            this.ed_textBox1.Label = "Keyword";
            this.ed_textBox1.MaxLength = 128;
            this.ed_textBox1.Name = "ed_textBox1";
            this.ed_textBox1.SizeString = "randomasstext";
            this.ed_textBox1.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // box3
            // 
            this.box3.Items.Add(this.label2);
            this.box3.Name = "box3";
            // 
            // label2
            // 
            this.label2.Label = "Destination";
            this.label2.Name = "label2";
            // 
            // box2
            // 
            this.box2.Items.Add(this.ed_colBox2);
            this.box2.Items.Add(this.ed_textBox2);
            this.box2.Name = "box2";
            // 
            // ed_colBox2
            // 
            this.ed_colBox2.Label = "Column";
            this.ed_colBox2.Name = "ed_colBox2";
            this.ed_colBox2.SizeString = "aaa";
            this.ed_colBox2.Text = null;
            // 
            // ed_textBox2
            // 
            this.ed_textBox2.Label = "Text";
            this.ed_textBox2.Name = "ed_textBox2";
            this.ed_textBox2.SizeString = "randomasstext";
            this.ed_textBox2.Text = null;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btn_go
            // 
            this.btn_go.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_go.Image = ((System.Drawing.Image)(resources.GetObject("btn_go.Image")));
            this.btn_go.Label = "Apply Autofill";
            this.btn_go.Name = "btn_go";
            this.btn_go.ScreenTip = "WARNING: Action cannot be undone!";
            this.btn_go.ShowImage = true;
            this.btn_go.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btn_go_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.box5);
            this.group2.Items.Add(this.box6);
            this.group2.Items.Add(this.separator3);
            this.group2.Items.Add(this.btn_go_signFlip);
            this.group2.Label = "Sign Flipper";
            this.group2.Name = "group2";
            // 
            // box5
            // 
            this.box5.Items.Add(this.label3);
            this.box5.Name = "box5";
            // 
            // label3
            // 
            this.label3.Label = "Source";
            this.label3.Name = "label3";
            // 
            // box6
            // 
            this.box6.Items.Add(this.ed_signFlip_colBox);
            this.box6.Items.Add(this.cb_pickSign);
            this.box6.Name = "box6";
            // 
            // ed_signFlip_colBox
            // 
            this.ed_signFlip_colBox.Label = "Column";
            this.ed_signFlip_colBox.Name = "ed_signFlip_colBox";
            this.ed_signFlip_colBox.SizeString = "aaa";
            this.ed_signFlip_colBox.Text = null;
            // 
            // cb_pickSign
            // 
            ribbonDropDownItem3.Label = "+ Positive";
            ribbonDropDownItem4.Label = "- Negative";
            this.cb_pickSign.Items.Add(ribbonDropDownItem3);
            this.cb_pickSign.Items.Add(ribbonDropDownItem4);
            this.cb_pickSign.Label = "Sign";
            this.cb_pickSign.Name = "cb_pickSign";
            this.cb_pickSign.SizeString = "all positive";
            this.cb_pickSign.Text = "+ Positive";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // btn_go_signFlip
            // 
            this.btn_go_signFlip.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_go_signFlip.Image = ((System.Drawing.Image)(resources.GetObject("btn_go_signFlip.Image")));
            this.btn_go_signFlip.Label = "Flip Signs";
            this.btn_go_signFlip.Name = "btn_go_signFlip";
            this.btn_go_signFlip.ShowImage = true;
            this.btn_go_signFlip.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btn_go_signFlip_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.box7);
            this.group3.Items.Add(this.box8);
            this.group3.Items.Add(this.separator4);
            this.group3.Items.Add(this.btn_go_cellCleanup);
            this.group3.Label = "Cell Cleanup";
            this.group3.Name = "group3";
            // 
            // box7
            // 
            this.box7.Items.Add(this.label4);
            this.box7.Name = "box7";
            // 
            // label4
            // 
            this.label4.Label = "Remove Characters";
            this.label4.Name = "label4";
            // 
            // box8
            // 
            this.box8.Items.Add(this.ed_cellCleanup_column);
            this.box8.Items.Add(this.ed_cellCleanup_characters);
            this.box8.Name = "box8";
            // 
            // ed_cellCleanup_column
            // 
            this.ed_cellCleanup_column.Label = "Column";
            this.ed_cellCleanup_column.Name = "ed_cellCleanup_column";
            this.ed_cellCleanup_column.SizeString = "aaa";
            this.ed_cellCleanup_column.Text = null;
            // 
            // ed_cellCleanup_characters
            // 
            this.ed_cellCleanup_characters.Label = "Characters";
            this.ed_cellCleanup_characters.Name = "ed_cellCleanup_characters";
            this.ed_cellCleanup_characters.SizeString = "aaaaaaaa";
            this.ed_cellCleanup_characters.SuperTip = "Enter characters to remove. Do not use spaces or commas to separate them.";
            this.ed_cellCleanup_characters.Text = null;
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // btn_go_cellCleanup
            // 
            this.btn_go_cellCleanup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_go_cellCleanup.Image = ((System.Drawing.Image)(resources.GetObject("btn_go_cellCleanup.Image")));
            this.btn_go_cellCleanup.Label = "Apply Cleanup";
            this.btn_go_cellCleanup.Name = "btn_go_cellCleanup";
            this.btn_go_cellCleanup.ShowImage = true;
            this.btn_go_cellCleanup.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btn_go_cellCleanup_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.box6.ResumeLayout(false);
            this.box6.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box7.ResumeLayout(false);
            this.box7.PerformLayout();
            this.box8.ResumeLayout(false);
            this.box8.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_colBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_textBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_colBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_textBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_go;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_go_signFlip;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box6;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_signFlip_colBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cb_pickSign;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box7;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box8;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_cellCleanup_column;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_go_cellCleanup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ed_cellCleanup_characters;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
