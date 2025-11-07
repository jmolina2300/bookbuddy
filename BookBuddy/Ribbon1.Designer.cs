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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem1 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem2 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
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
            this.group1.Items.Add(this.btn_go);
            this.group1.Label = "Transaction Description Utilities";
            this.group1.Name = "group1";

            // 
            // btn_go
            // 
            this.btn_go.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_go.Image = ((System.Drawing.Image)(resources.GetObject("btn_go.Image")));
            this.btn_go.Label = "Launch Autofiller Dialog";
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
            ribbonDropDownItem1.Label = "+ Positive";
            ribbonDropDownItem2.Label = "- Negative";
            this.cb_pickSign.Items.Add(ribbonDropDownItem1);
            this.cb_pickSign.Items.Add(ribbonDropDownItem2);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_go;
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
