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
            this.grMarkdown = this.Factory.CreateRibbonGroup();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.tabMarkdown.SuspendLayout();
            this.grMarkdown.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMarkdown
            // 
            this.tabMarkdown.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMarkdown.Groups.Add(this.grMarkdown);
            this.tabMarkdown.Label = "Markdown";
            this.tabMarkdown.Name = "tabMarkdown";
            // 
            // grMarkdown
            // 
            this.grMarkdown.Items.Add(this.btnSave);
            this.grMarkdown.Label = "Markdown";
            this.grMarkdown.Name = "grMarkdown";
            // 
            // btnSave
            // 
            this.btnSave.Label = "Сохранить";
            this.btnSave.Name = "btnSave";
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // MarkdownRibbon
            // 
            this.Name = "MarkdownRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabMarkdown);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MarkdownRibbon_Load);
            this.tabMarkdown.ResumeLayout(false);
            this.tabMarkdown.PerformLayout();
            this.grMarkdown.ResumeLayout(false);
            this.grMarkdown.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
    }

    partial class ThisRibbonCollection
    {
        internal MarkdownRibbon MarkdownRibbon
        {
            get { return this.GetRibbon<MarkdownRibbon>(); }
        }
    }
}
