namespace BoggleMaker {
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.MainRibbonGrp = this.Factory.CreateRibbonGroup();
            this.InsertNew4 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.MainRibbonGrp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.MainRibbonGrp);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // MainRibbonGrp
            // 
            this.MainRibbonGrp.Items.Add(this.InsertNew4);
            this.MainRibbonGrp.Label = "BoggleMaker";
            this.MainRibbonGrp.Name = "MainRibbonGrp";
            // 
            // InsertNew4
            // 
            this.InsertNew4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.InsertNew4.Image = ((System.Drawing.Image)(resources.GetObject("InsertNew4.Image")));
            this.InsertNew4.ImageName = "grid4";
            this.InsertNew4.Label = "New 4x4 Boggle";
            this.InsertNew4.Name = "InsertNew4";
            this.InsertNew4.ShowImage = true;
            this.InsertNew4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertNew4_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.MainRibbonGrp.ResumeLayout(false);
            this.MainRibbonGrp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup MainRibbonGrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertNew4;
    }

    partial class ThisRibbonCollection {
        internal MainRibbon Ribbon1 {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
