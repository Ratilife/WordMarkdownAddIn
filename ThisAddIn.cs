using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;





namespace WordMarkdownAddIn
{
    public partial class ThisAddIn
    {
        public static CustomTaskPane MarkdownPane { get; private set; }
        public static Controls.TaskPaneControl PaneControl { get; private set; }
        public static MarkdownRibbon Ribbon { get; private set; }
        public Dictionary<string, object> Properties { get; private set; }
        public static ThisAddIn Instance { get; private set; }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 1. Инициализируем словарь свойств
            Properties = new Dictionary<string, object>();

            // 2. Сохраняем экземпляр для глобального доступа
            Instance = this;

            // 3. Создаем и настраиваем панель Markdown-редактора
            PaneControl = new Controls.TaskPaneControl();
            MarkdownPane = this.CustomTaskPanes.Add(PaneControl, "Markdown");
            MarkdownPane.Visible = true;
            MarkdownPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            MarkdownPane.Width = 600;
            // 4. Загружаем сохраненный Markdown из текущего документа
            try
            {
                var md = Services.DocumentSyncService.LoadMarkdownFromActiveDocument(Application);
                if (!string.IsNullOrEmpty(md))
                {
                    PaneControl.SetMarkdown(md);
                }
            }
            catch { /* Игнорируем ошибки при старте */ }


            // 5. Подписываемся на события Word для автоматического сохранения
            // Отследить сохранение Word, чтобы сохранить текущую уценку в CustomXMLPart
            try
            {
                this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            }
            catch { }

            // 6. Инициализируем ленту Ribbon
            try
            {
                Ribbon = new MarkdownRibbon();
            }
            catch { }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 1. Отписываемся от событий Word
            try
            {
                this.Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
                
            }
            catch { /* Игнорируем ошибки отписки */ }

            // 2. Сохраняем настройки пользователя
            try
            {
                
                    this.Properties["PaneWidth"] = MarkdownPane.Width;
                    this.Properties["PaneVisible"] = MarkdownPane.Visible;
                

            }
            catch { /* Игнорируем ошибки сохранения */ }

            // 3. Освобождаем COM-объекты
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Application);
            }
            catch { /* Игнорируем ошибки освобождения */ }
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                var md = PaneControl.GetCachedMarkdown() ?? string.Empty;
                Services.DocumentSyncService.SaveMarkdownToActiveDocument(Application, md);
            }
            catch { }
        }

      

        public void TogglePane()
        {
            if (MarkdownPane != null)
            {
                MarkdownPane.Visible = !MarkdownPane.Visible;
            }
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
