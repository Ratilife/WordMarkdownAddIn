namespace WordMarkdownAddIn
{
    partial class RibbonMarkdown : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMarkdown()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grMarkdown = this.Factory.CreateRibbonGroup();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grMarkdown.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grMarkdown);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
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
            // RibbonMarkdown
            // 
            this.Name = "RibbonMarkdown";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grMarkdown.ResumeLayout(false);
            this.grMarkdown.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMarkdown Ribbon1
        {
            get { return this.GetRibbon<RibbonMarkdown>(); }
        }
    }
}
