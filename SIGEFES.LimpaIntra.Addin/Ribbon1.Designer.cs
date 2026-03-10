namespace SIGEFES.LimpaIntra.Addin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Auditoria = this.Factory.CreateRibbonGroup();
            this.BtnLimparIntra = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Auditoria.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.Auditoria);
            this.tab1.Label = "SIGEFES";
            this.tab1.Name = "tab1";
            // 
            // Auditoria
            // 
            this.Auditoria.Items.Add(this.BtnLimparIntra);
            this.Auditoria.Label = "Auditoria";
            this.Auditoria.Name = "Auditoria";
            // 
            // BtnLimparIntra
            // 
            this.BtnLimparIntra.Label = "Limpar intraoffs";
            this.BtnLimparIntra.Name = "BtnLimparIntra";
            this.BtnLimparIntra.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLimparIntra_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Auditoria.ResumeLayout(false);
            this.Auditoria.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Auditoria;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLimparIntra;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
