namespace WordMarkdownAddIn
{
    partial class MarkdownRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MarkdownRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabMarkdown = this.Factory.CreateRibbonTab();
            this.grpFile = this.Factory.CreateRibbonGroup();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.btnPanel = this.Factory.CreateRibbonButton();
            this.btnOpen = this.Factory.CreateRibbonButton();
            this.grpFormat = this.Factory.CreateRibbonGroup();
            this.bBold = this.Factory.CreateRibbonButton();
            this.tabMarkdown.SuspendLayout();
            this.grpFile.SuspendLayout();
            this.grpFormat.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMarkdown
            // 
            this.tabMarkdown.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMarkdown.Groups.Add(this.grpFile);
            this.tabMarkdown.Groups.Add(this.grpFormat);
            this.tabMarkdown.Label = "Markdown";
            this.tabMarkdown.Name = "tabMarkdown";
            // 
            // grpFile
            // 
            this.grpFile.Items.Add(this.btnSave);
            this.grpFile.Items.Add(this.btnPanel);
            this.grpFile.Items.Add(this.btnOpen);
            this.grpFile.Label = "Файл";
            this.grpFile.Name = "grpFile";
            // 
            // btnSave
            // 
            this.btnSave.Label = "Сохранить .md";
            this.btnSave.Name = "btnSave";
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // btnPanel
            // 
            this.btnPanel.Label = "Панель";
            this.btnPanel.Name = "btnPanel";
            this.btnPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPanel_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.Label = "Открыть .md";
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpen_Click);
            // 
            // grpFormat
            // 
            this.grpFormat.Items.Add(this.bBold);
            this.grpFormat.Label = "Форматирование";
            this.grpFormat.Name = "grpFormat";
            // 
            // bBold
            // 
            this.bBold.Label = "Жирный";
            this.bBold.Name = "bBold";
            this.bBold.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bBold_Click);
            // 
            // MarkdownRibbon
            // 
            this.Name = "MarkdownRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabMarkdown);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MarkdownRibbon_Load);
            this.tabMarkdown.ResumeLayout(false);
            this.tabMarkdown.PerformLayout();
            this.grpFile.ResumeLayout(false);
            this.grpFile.PerformLayout();
            this.grpFormat.ResumeLayout(false);
            this.grpFormat.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPanel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpen;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bBold;
    }

    partial class ThisRibbonCollection
    {
        internal MarkdownRibbon MarkdownRibbon
        {
            get { return this.GetRibbon<MarkdownRibbon>(); }
        }
    }
}
