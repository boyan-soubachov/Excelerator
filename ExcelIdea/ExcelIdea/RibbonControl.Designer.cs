namespace ExcelIdea
{
    partial class RibbonControl : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonControl()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.ddl_InteropMethod = this.Factory.CreateRibbonDropDown();
            this.btn_start = this.Factory.CreateRibbonButton();
            this.mnu_Tools = this.Factory.CreateRibbonMenu();
            this.btn_Tools_Settings = this.Factory.CreateRibbonButton();
            this.chkOpenWhenDone = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_start);
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.mnu_Tools);
            this.group1.Label = "ExcelIdea";
            this.group1.Name = "group1";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.ddl_InteropMethod);
            this.box1.Items.Add(this.chkOpenWhenDone);
            this.box1.Name = "box1";
            // 
            // ddl_InteropMethod
            // 
            ribbonDropDownItemImpl1.Label = "File";
            ribbonDropDownItemImpl1.ScreenTip = "File";
            ribbonDropDownItemImpl1.SuperTip = "Transfer the data by saving and loading an Excel file";
            ribbonDropDownItemImpl1.Tag = "file";
            ribbonDropDownItemImpl2.Label = "IPC";
            ribbonDropDownItemImpl2.ScreenTip = "Inter-process Communication";
            ribbonDropDownItemImpl2.SuperTip = "Transfer the data invisibly using IPC";
            ribbonDropDownItemImpl2.Tag = "ipc";
            this.ddl_InteropMethod.Items.Add(ribbonDropDownItemImpl1);
            this.ddl_InteropMethod.Items.Add(ribbonDropDownItemImpl2);
            this.ddl_InteropMethod.Label = "Comm. Method:";
            this.ddl_InteropMethod.Name = "ddl_InteropMethod";
            this.ddl_InteropMethod.ScreenTip = "Excel <-> ACE communication method";
            this.ddl_InteropMethod.SuperTip = "Specifies the method used for transferring Excel data to and from the Advanced Co" +
    "mputation Engine";
            this.ddl_InteropMethod.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddl_InteropMethod_SelectionChanged);
            // 
            // btn_start
            // 
            this.btn_start.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_start.Label = "Start";
            this.btn_start.Name = "btn_start";
            this.btn_start.OfficeImageId = "StartTimer";
            this.btn_start.ShowImage = true;
            this.btn_start.SuperTip = "Start the workbook calculation";
            this.btn_start.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_start_Click);
            // 
            // mnu_Tools
            // 
            this.mnu_Tools.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mnu_Tools.Items.Add(this.btn_Tools_Settings);
            this.mnu_Tools.Label = "Tools";
            this.mnu_Tools.Name = "mnu_Tools";
            this.mnu_Tools.ShowImage = true;
            // 
            // btn_Tools_Settings
            // 
            this.btn_Tools_Settings.Label = "Settings";
            this.btn_Tools_Settings.Name = "btn_Tools_Settings";
            this.btn_Tools_Settings.ShowImage = true;
            this.btn_Tools_Settings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Tools_Settings_Click);
            // 
            // chkOpenWhenDone
            // 
            this.chkOpenWhenDone.Label = "Open upon completion";
            this.chkOpenWhenDone.Name = "chkOpenWhenDone";
            // 
            // RibbonControl
            // 
            this.Name = "RibbonControl";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonControl_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_start;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnu_Tools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Tools_Settings;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddl_InteropMethod;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkOpenWhenDone;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonControl RibbonControl
        {
            get { return this.GetRibbon<RibbonControl>(); }
        }
    }
}
