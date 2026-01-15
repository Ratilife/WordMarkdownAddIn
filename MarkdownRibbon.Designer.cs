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
            this.bItalic = this.Factory.CreateRibbonButton();
            this.bStrike = this.Factory.CreateRibbonButton();
            this.bCode = this.Factory.CreateRibbonButton();
            this.grpInsert = this.Factory.CreateRibbonGroup();
            this.bH1 = this.Factory.CreateRibbonButton();
            this.bH2 = this.Factory.CreateRibbonButton();
            this.bH3 = this.Factory.CreateRibbonButton();
            this.bList = this.Factory.CreateRibbonButton();
            this.bNumList = this.Factory.CreateRibbonButton();
            this.bCheckbox = this.Factory.CreateRibbonButton();
            this.bCheckboxTrue = this.Factory.CreateRibbonButton();
            this.bTable = this.Factory.CreateRibbonButton();
            this.bLink = this.Factory.CreateRibbonButton();
            this.bImage = this.Factory.CreateRibbonButton();
            this.bHR = this.Factory.CreateRibbonButton();
            this.bCodeBlock = this.Factory.CreateRibbonButton();
            this.bMermaid = this.Factory.CreateRibbonButton();
            this.bMath = this.Factory.CreateRibbonButton();
            this.grpConvert = this.Factory.CreateRibbonGroup();
            this.btnConvert = this.Factory.CreateRibbonButton();
            this.btnConvertMD_DocNotF = this.Factory.CreateRibbonButton();
            this.btnConvertMD_Doc = this.Factory.CreateRibbonButton();
            this.btnFormatMarkdown = this.Factory.CreateRibbonButton();
            this.grpExport = this.Factory.CreateRibbonGroup();
            this.btnExportMermaid = this.Factory.CreateRibbonButton();
            this.tabMarkdown.SuspendLayout();
            this.grpFile.SuspendLayout();
            this.grpFormat.SuspendLayout();
            this.grpInsert.SuspendLayout();
            this.grpConvert.SuspendLayout();
            this.grpExport.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMarkdown
            // 
            this.tabMarkdown.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMarkdown.Groups.Add(this.grpFile);
            this.tabMarkdown.Groups.Add(this.grpFormat);
            this.tabMarkdown.Groups.Add(this.grpInsert);
            this.tabMarkdown.Groups.Add(this.grpConvert);
            this.tabMarkdown.Groups.Add(this.grpExport);
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
            this.grpFormat.Items.Add(this.bItalic);
            this.grpFormat.Items.Add(this.bStrike);
            this.grpFormat.Items.Add(this.bCode);
            this.grpFormat.Label = "Форматирование";
            this.grpFormat.Name = "grpFormat";
            // 
            // bBold
            // 
            this.bBold.Label = "Жирный";
            this.bBold.Name = "bBold";
            this.bBold.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bBold_Click);
            // 
            // bItalic
            // 
            this.bItalic.Label = "Курсив";
            this.bItalic.Name = "bItalic";
            this.bItalic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bItalic_Click);
            // 
            // bStrike
            // 
            this.bStrike.Label = "Зачеркнуть";
            this.bStrike.Name = "bStrike";
            this.bStrike.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bStrike_Click);
            // 
            // bCode
            // 
            this.bCode.Label = "Код";
            this.bCode.Name = "bCode";
            this.bCode.Visible = false;
            // 
            // grpInsert
            // 
            this.grpInsert.Items.Add(this.bH1);
            this.grpInsert.Items.Add(this.bH2);
            this.grpInsert.Items.Add(this.bH3);
            this.grpInsert.Items.Add(this.bList);
            this.grpInsert.Items.Add(this.bNumList);
            this.grpInsert.Items.Add(this.bCheckbox);
            this.grpInsert.Items.Add(this.bCheckboxTrue);
            this.grpInsert.Items.Add(this.bTable);
            this.grpInsert.Items.Add(this.bLink);
            this.grpInsert.Items.Add(this.bImage);
            this.grpInsert.Items.Add(this.bHR);
            this.grpInsert.Items.Add(this.bCodeBlock);
            this.grpInsert.Items.Add(this.bMermaid);
            this.grpInsert.Items.Add(this.bMath);
            this.grpInsert.Label = "Вставка";
            this.grpInsert.Name = "grpInsert";
            // 
            // bH1
            // 
            this.bH1.Label = "H1";
            this.bH1.Name = "bH1";
            this.bH1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bH1_Click);
            // 
            // bH2
            // 
            this.bH2.Label = "H2";
            this.bH2.Name = "bH2";
            this.bH2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bH2_Click);
            // 
            // bH3
            // 
            this.bH3.Label = "H3";
            this.bH3.Name = "bH3";
            this.bH3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bH3_Click);
            // 
            // bList
            // 
            this.bList.Label = "Список -";
            this.bList.Name = "bList";
            this.bList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bList_Click);
            // 
            // bNumList
            // 
            this.bNumList.Label = "Список 1.";
            this.bNumList.Name = "bNumList";
            this.bNumList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bNumList_Click);
            // 
            // bCheckbox
            // 
            this.bCheckbox.Label = "Чекбокс(Ложь)";
            this.bCheckbox.Name = "bCheckbox";
            this.bCheckbox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bCheckbox_Click);
            // 
            // bCheckboxTrue
            // 
            this.bCheckboxTrue.Label = "Чекбокс(Истина)";
            this.bCheckboxTrue.Name = "bCheckboxTrue";
            this.bCheckboxTrue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bCheckboxTrue_Click);
            // 
            // bTable
            // 
            this.bTable.Label = "Таблица";
            this.bTable.Name = "bTable";
            this.bTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bTable_Click);
            // 
            // bLink
            // 
            this.bLink.Label = "Ссылка";
            this.bLink.Name = "bLink";
            this.bLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bLink_Click);
            // 
            // bImage
            // 
            this.bImage.Enabled = false;
            this.bImage.Label = "Изображение";
            this.bImage.Name = "bImage";
            this.bImage.Visible = false;
            this.bImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bImage_Click);
            // 
            // bHR
            // 
            this.bHR.Label = "Разделитель";
            this.bHR.Name = "bHR";
            this.bHR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bHR_Click);
            // 
            // bCodeBlock
            // 
            this.bCodeBlock.Label = "Код-блок";
            this.bCodeBlock.Name = "bCodeBlock";
            this.bCodeBlock.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bCodeBlock_Click);
            // 
            // bMermaid
            // 
            this.bMermaid.Label = "Mermaid";
            this.bMermaid.Name = "bMermaid";
            this.bMermaid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bMermaid_Click);
            // 
            // bMath
            // 
            this.bMath.Label = "Формула";
            this.bMath.Name = "bMath";
            this.bMath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bMath_Click);
            // 
            // grpConvert
            // 
            this.grpConvert.Items.Add(this.btnConvert);
            this.grpConvert.Items.Add(this.btnConvertMD_DocNotF);
            this.grpConvert.Items.Add(this.btnConvertMD_Doc);
            this.grpConvert.Items.Add(this.btnFormatMarkdown);
            this.grpConvert.Label = "Преобразование";
            this.grpConvert.Name = "grpConvert";
            // 
            // btnConvert
            // 
            this.btnConvert.Label = "Word → Markdown";
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvert_Click);
            // 
            // btnConvertMD_DocNotF
            // 
            this.btnConvertMD_DocNotF.Label = "Word → Markdown (без форматирования)";
            this.btnConvertMD_DocNotF.Name = "btnConvertMD_DocNotF";
            this.btnConvertMD_DocNotF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvertMD_DocNotF_Click);
            // 
            // btnConvertMD_Doc
            // 
            this.btnConvertMD_Doc.Label = "Markdown → Word";
            this.btnConvertMD_Doc.Name = "btnConvertMD_Doc";
            this.btnConvertMD_Doc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyToWord_Click);
            // 
            // btnFormatMarkdown
            // 
            this.btnFormatMarkdown.Label = "Форматировать Markdown";
            this.btnFormatMarkdown.Name = "btnFormatMarkdown";
            this.btnFormatMarkdown.Visible = false;
            this.btnFormatMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatMarkdown_Click);
            // 
            // grpExport
            // 
            this.grpExport.Items.Add(this.btnExportMermaid);
            this.grpExport.Label = "Экспорт";
            this.grpExport.Name = "grpExport";
            // 
            // btnExportMermaid
            // 
            this.btnExportMermaid.Label = "Экспорт Mermaid в PNG";
            this.btnExportMermaid.Name = "btnExportMermaid";
            this.btnExportMermaid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportMermaid_Click);
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
            this.grpInsert.ResumeLayout(false);
            this.grpInsert.PerformLayout();
            this.grpConvert.ResumeLayout(false);
            this.grpConvert.PerformLayout();
            this.grpExport.ResumeLayout(false);
            this.grpExport.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bItalic;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bStrike;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInsert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bH1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bH2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bH3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bNumList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bCheckbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bLink;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bHR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bCodeBlock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bMermaid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bMath;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertMD_Doc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertMD_DocNotF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bCheckboxTrue;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportMermaid;
    }

    partial class ThisRibbonCollection
    {
        internal MarkdownRibbon MarkdownRibbon
        {
            get { return this.GetRibbon<MarkdownRibbon>(); }
        }
    }
}
