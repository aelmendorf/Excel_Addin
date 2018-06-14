namespace Epi_AddIn {
    partial class Epi_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Epi_Ribbon()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonBox box1;
            this.epi_tab_1 = this.Factory.CreateRibbonTab();
            this.gatherData = this.Factory.CreateRibbonGroup();
            this.commonFiles = this.Factory.CreateRibbonGroup();
            this.importBurnIn = this.Factory.CreateRibbonGroup();
            this.getSpectrum = this.Factory.CreateRibbonButton();
            this.openEWAT = this.Factory.CreateRibbonButton();
            this.importBurn = this.Factory.CreateRibbonButton();
            this.testType = this.Factory.CreateRibbonComboBox();
            box1 = this.Factory.CreateRibbonBox();
            this.epi_tab_1.SuspendLayout();
            this.gatherData.SuspendLayout();
            this.commonFiles.SuspendLayout();
            this.importBurnIn.SuspendLayout();
            box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // epi_tab_1
            // 
            this.epi_tab_1.Groups.Add(this.gatherData);
            this.epi_tab_1.Groups.Add(this.commonFiles);
            this.epi_tab_1.Groups.Add(this.importBurnIn);
            this.epi_tab_1.Label = "Epi Add-In Tab";
            this.epi_tab_1.Name = "epi_tab_1";
            // 
            // gatherData
            // 
            this.gatherData.Items.Add(this.getSpectrum);
            this.gatherData.Label = "Gather Data";
            this.gatherData.Name = "gatherData";
            // 
            // commonFiles
            // 
            this.commonFiles.Items.Add(this.openEWAT);
            this.commonFiles.Label = "Open Common Files";
            this.commonFiles.Name = "commonFiles";
            // 
            // importBurnIn
            // 
            this.importBurnIn.Items.Add(box1);
            this.importBurnIn.Items.Add(this.testType);
            this.importBurnIn.Label = "Input Burn-In";
            this.importBurnIn.Name = "importBurnIn";
            // 
            // getSpectrum
            // 
            this.getSpectrum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.getSpectrum.Image = global::Epi_AddIn.Properties.Resources.masking_1;
            this.getSpectrum.Label = "Get Spectrum Data";
            this.getSpectrum.Name = "getSpectrum";
            this.getSpectrum.ShowImage = true;
            this.getSpectrum.SuperTip = "Select wafers then press";
            this.getSpectrum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.getSpectrum_Click);
            // 
            // openEWAT
            // 
            this.openEWAT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.openEWAT.Image = global::Epi_AddIn.Properties.Resources.data_protection;
            this.openEWAT.Label = "Open EWAT";
            this.openEWAT.Name = "openEWAT";
            this.openEWAT.ShowImage = true;
            this.openEWAT.SuperTip = "Opens EWAT Database";
            this.openEWAT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openEWAT_Click);
            // 
            // importBurn
            // 
            this.importBurn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.importBurn.Image = global::Epi_AddIn.Properties.Resources.data_clipart_free_business_clipart_collection;
            this.importBurn.Label = "Import Data";
            this.importBurn.Name = "importBurn";
            this.importBurn.ShowImage = true;
            this.importBurn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importBurn_Click);
            // 
            // testType
            // 
            this.testType.Image = global::Epi_AddIn.Properties.Resources.sign_check_icon;
            ribbonDropDownItemImpl1.Label = "Initial";
            ribbonDropDownItemImpl2.Label = "After";
            this.testType.Items.Add(ribbonDropDownItemImpl1);
            this.testType.Items.Add(ribbonDropDownItemImpl2);
            this.testType.Label = "Test Type";
            this.testType.Name = "testType";
            this.testType.ShowImage = true;
            this.testType.Text = null;
            // 
            // box1
            // 
            box1.Items.Add(this.importBurn);
            box1.Name = "box1";
            // 
            // Epi_Ribbon
            // 
            this.Name = "Epi_Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.epi_tab_1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Epi_Ribbon_Load);
            this.epi_tab_1.ResumeLayout(false);
            this.epi_tab_1.PerformLayout();
            this.gatherData.ResumeLayout(false);
            this.gatherData.PerformLayout();
            this.commonFiles.ResumeLayout(false);
            this.commonFiles.PerformLayout();
            this.importBurnIn.ResumeLayout(false);
            this.importBurnIn.PerformLayout();
            box1.ResumeLayout(false);
            box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab epi_tab_1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gatherData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton getSpectrum;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup commonFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openEWAT;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup importBurnIn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importBurn;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox testType;
    }

    partial class ThisRibbonCollection {
        internal Epi_Ribbon Epi_Ribbon {
            get { return this.GetRibbon<Epi_Ribbon>(); }
        }
    }
}
