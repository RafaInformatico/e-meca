namespace EMecaAddin
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabEMeca = this.Factory.CreateRibbonTab();
            this.groupTest = this.Factory.CreateRibbonGroup();
            this.btnOpenModel = this.Factory.CreateRibbonButton();
            this.ddTime = this.Factory.CreateRibbonDropDown();
            this.btnStart = this.Factory.CreateRibbonButton();
            this.btnStop = this.Factory.CreateRibbonButton();
            this.groupTypingData = this.Factory.CreateRibbonGroup();
            this.btnTypingData = this.Factory.CreateRibbonButton();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.timerTest = new System.Windows.Forms.Timer(this.components);
            this.btnNewTest = this.Factory.CreateRibbonButton();
            this.tabEMeca.SuspendLayout();
            this.groupTest.SuspendLayout();
            this.groupTypingData.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEMeca
            // 
            this.tabEMeca.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEMeca.Groups.Add(this.groupTest);
            this.tabEMeca.Groups.Add(this.groupTypingData);
            this.tabEMeca.Label = "eMeca";
            this.tabEMeca.Name = "tabEMeca";
            // 
            // groupTest
            // 
            this.groupTest.Items.Add(this.btnNewTest);
            this.groupTest.Items.Add(this.btnOpenModel);
            this.groupTest.Items.Add(this.ddTime);
            this.groupTest.Items.Add(this.btnStart);
            this.groupTest.Items.Add(this.btnStop);
            this.groupTest.Label = "Ejercicio";
            this.groupTest.Name = "groupTest";
            // 
            // btnOpenModel
            // 
            this.btnOpenModel.Label = "Abrir modelo";
            this.btnOpenModel.Name = "btnOpenModel";
            this.btnOpenModel.OfficeImageId = "FileOpen";
            this.btnOpenModel.ShowImage = true;
            this.btnOpenModel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnOpenModel_Click);
            // 
            // ddTime
            // 
            this.ddTime.Label = "Tiempo";
            this.ddTime.Name = "ddTime";
            // 
            // btnStart
            // 
            this.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStart.Label = "Empezar";
            this.btnStart.Name = "btnStart";
            this.btnStart.OfficeImageId = "RecordingPlay";
            this.btnStart.ShowImage = true;
            this.btnStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnStart_Click);
            // 
            // btnStop
            // 
            this.btnStop.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStop.Label = "Terminar";
            this.btnStop.Name = "btnStop";
            this.btnStop.OfficeImageId = "RecordingStop";
            this.btnStop.ShowImage = true;
            this.btnStop.Visible = false;
            this.btnStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnStop_Click);
            // 
            // groupTypingData
            // 
            this.groupTypingData.Items.Add(this.btnTypingData);
            this.groupTypingData.Label = "Pulsaciones";
            this.groupTypingData.Name = "groupTypingData";
            // 
            // btnTypingData
            // 
            this.btnTypingData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTypingData.Image = global::EMecaAddin.Properties.Resources.CheckText;
            this.btnTypingData.Label = "Resultado";
            this.btnTypingData.Name = "btnTypingData";
            this.btnTypingData.ShowImage = true;
            this.btnTypingData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetTypingData_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "model";
            this.openFileDialog.Filter = "Word|*.doc;*.docx";
            // 
            // timerTest
            // 
            this.timerTest.Tick += new System.EventHandler(this.TimerTest_Tick);
            // 
            // btnNewTest
            // 
            this.btnNewTest.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewTest.Label = "Nuevo";
            this.btnNewTest.Name = "btnNewTest";
            this.btnNewTest.OfficeImageId = "Refresh";
            this.btnNewTest.ShowImage = true;
            this.btnNewTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewTest_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabEMeca);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabEMeca.ResumeLayout(false);
            this.tabEMeca.PerformLayout();
            this.groupTest.ResumeLayout(false);
            this.groupTest.PerformLayout();
            this.groupTypingData.ResumeLayout(false);
            this.groupTypingData.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEMeca;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTypingData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTypingData;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenModel;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStop;
        private System.Windows.Forms.Timer timerTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewTest;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
