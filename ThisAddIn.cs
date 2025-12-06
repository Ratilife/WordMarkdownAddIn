using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using WordMarkdownAddIn.Properties;





namespace WordMarkdownAddIn
{
    public partial class ThisAddIn
    {
        // Словари для хранения панелей для каждого окна документа
        private static Dictionary<Word.Window, CustomTaskPane> _markdownPanes = new Dictionary<Word.Window, CustomTaskPane>();
        private static Dictionary<Word.Window, Controls.TaskPaneControl> _paneControls = new Dictionary<Word.Window, Controls.TaskPaneControl>();

        // Статические свойства для обратной совместимости - возвращают панель для активного документа
        public static CustomTaskPane MarkdownPane 
        { 
            get 
            {
                if (Instance?.Application?.ActiveWindow != null)
                {
                    _markdownPanes.TryGetValue(Instance.Application.ActiveWindow, out var pane);
                    return pane;
                }
                return null;
            } 
        }

        public static Controls.TaskPaneControl PaneControl 
        { 
            get 
            {
                if (Instance?.Application?.ActiveWindow != null)
                {
                    _paneControls.TryGetValue(Instance.Application.ActiveWindow, out var control);
                    return control;
                }
                return null;
            } 
        }

        public static MarkdownRibbon Ribbon { get; private set; }
        public Dictionary<string, object> Properties { get; private set; }
        public static ThisAddIn Instance { get; private set; }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 1. Инициализируем словарь свойств
            Properties = new Dictionary<string, object>();

            // 2. Сохраняем экземпляр для глобального доступа
            Instance = this;

            // 3. Подписываемся на события Word для управления панелями
            try
            {
                this.Application.DocumentOpen += Application_DocumentOpen;
                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
                this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            }
            catch { }

            // 4. Создаем панель для активного документа (если он есть)
            try
            {
                if (this.Application.ActiveWindow != null)
                {
                    EnsurePaneForWindow(this.Application.ActiveWindow);
                }
            }
            catch { /* Игнорируем ошибки при старте */ }

            // 5. Инициализируем ленту Ribbon
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
                this.Application.DocumentOpen -= Application_DocumentOpen;
                this.Application.WindowActivate -= Application_WindowActivate;
                this.Application.DocumentBeforeClose -= Application_DocumentBeforeClose;
                this.Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
                
            }
            catch { /* Игнорируем ошибки отписки */ }

            // 2. Сохраняем настройки пользователя
            try
            {
                
                    if (_markdownPanes.Count > 0)
                    {
                        var firstPane = _markdownPanes.Values.First();
                        this.Properties["PaneWidth"] = firstPane.Width;
                        this.Properties["PaneVisible"] = firstPane.Visible;
                    }
                

            }
            catch { /* Игнорируем ошибки сохранения */ }

            // 3. Сохраняем Markdown для всех открытых документов
            try
            {
                foreach (var kvp in _paneControls.ToList())
                {
                    try
                    {
                        var doc = kvp.Key.Document;
                        if (doc != null)
                        {
                            var md = kvp.Value.GetCachedMarkdown() ?? string.Empty;
                            SaveMarkdownToDocument(doc, md);
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // 4. Удаляем все панели из коллекции CustomTaskPanes
            try
            {
                foreach (var pane in _markdownPanes.Values.ToList())
                {
                    try
                    {
                        this.CustomTaskPanes.Remove(pane);
                    }
                    catch { }
                }
            }
            catch { }

            // 5. Очищаем словари
            _markdownPanes.Clear();
            _paneControls.Clear();

            // 6. Освобождаем COM-объекты
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Application);
            }
            catch { /* Игнорируем ошибки освобождения */ }
        }

        // Создает панель для указанного окна, если её еще нет
        private void EnsurePaneForWindow(Word.Window window)
        {
            if (window == null) return;

            // Проверяем, есть ли уже панель для этого окна
            if (_markdownPanes.ContainsKey(window))
                return;

            try
            {
                // Создаем новый контрол и панель
                var paneControl = new Controls.TaskPaneControl();
                var pane = this.CustomTaskPanes.Add(paneControl, "Markdown", window);
                
                // Настраиваем панель
                pane.Visible = true;
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                
                // Восстанавливаем сохраненную ширину, если есть
                if (Properties.ContainsKey("PaneWidth"))
                {
                    pane.Width = (int)Properties["PaneWidth"];
                }
                else
                {
                    pane.Width = 600;
                }

                // Восстанавливаем видимость, если сохранена
                if (Properties.ContainsKey("PaneVisible"))
                {
                    pane.Visible = (bool)Properties["PaneVisible"];
                }

                // Сохраняем в словари
                _markdownPanes[window] = pane;
                _paneControls[window] = paneControl;

                // Загружаем сохраненный Markdown из документа
                try
                {
                    var doc = window.Document;
                    var md = LoadMarkdownFromDocument(doc);
                    if (!string.IsNullOrEmpty(md))
                    {
                        paneControl.SetMarkdown(md);
                    }
                }
                catch { /* Игнорируем ошибки загрузки */ }

                // Восстанавливаем сохраненный режим отображения
                try
                {
                    var savedMode = Settings.Default.ViewMode;
                    if (!string.IsNullOrEmpty(savedMode) && (savedMode == "split" || savedMode == "markdown" || savedMode == "html"))
                    {
                        // Небольшая задержка для инициализации WebView2
                        System.Threading.Tasks.Task.Delay(500).ContinueWith(t =>
                        {
                            System.Windows.Forms.Application.DoEvents();
                            paneControl.SetViewMode(savedMode);
                        });
                    }
                }
                catch { /* Игнорируем ошибки при загрузке режима */ }
            }
            catch { /* Игнорируем ошибки создания панели */ }
        }

        // Обработчик открытия документа
        private void Application_DocumentOpen(Word.Document Doc)
        {
            try
            {
                // Находим окно документа и создаем для него панель
                if (Doc.Windows.Count > 0)
                {
                    var window = Doc.Windows[1]; // Берем первое окно документа
                    EnsurePaneForWindow(window);
                }
            }
            catch { }
        }

        // Обработчик активации окна
        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            try
            {
                // Создаем панель для активированного окна, если её нет
                if (Wn != null)
                {
                    EnsurePaneForWindow(Wn);
                }
            }
            catch { }
        }

        // Обработчик закрытия документа
        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            try
            {
                // Сохраняем Markdown перед закрытием документа
                Word.Window windowToRemove = null;
                foreach (var kvp in _paneControls)
                {
                    if (kvp.Key.Document == Doc)
                    {
                        windowToRemove = kvp.Key;
                        var paneControl = kvp.Value;
                        try
                        {
                            var md = paneControl.GetCachedMarkdown() ?? string.Empty;
                            SaveMarkdownToDocument(Doc, md);
                        }
                        catch { }

                        // Удаляем панель из коллекции CustomTaskPanes
                        if (_markdownPanes.ContainsKey(windowToRemove))
                        {
                            var pane = _markdownPanes[windowToRemove];
                            try
                            {
                                this.CustomTaskPanes.Remove(pane);
                            }
                            catch { }
                        }
                        break;
                    }
                }

                // Удаляем из словарей
                if (windowToRemove != null)
                {
                    _markdownPanes.Remove(windowToRemove);
                    _paneControls.Remove(windowToRemove);
                }
            }
            catch { }
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                // Находим панель для этого документа и сохраняем Markdown
                foreach (var kvp in _paneControls)
                {
                    if (kvp.Key.Document == Doc)
                    {
                        var md = kvp.Value.GetCachedMarkdown() ?? string.Empty;
                        SaveMarkdownToDocument(Doc, md);
                        break;
                    }
                }
            }
            catch { }
        }

      

        public void TogglePane()
        {
            var pane = MarkdownPane;
            if (pane != null)
            {
                pane.Visible = !pane.Visible;
            }
        }

        // Вспомогательные методы для работы с Markdown конкретного документа
        private string LoadMarkdownFromDocument(Word.Document doc)
        {
            if (doc == null) return null;
            try
            {
                var part = FindExistingPart(doc);
                if (part == null) return null;
                var node = part.SelectSingleNode("/*[local-name()='markdown']/*[local-name()='content']");
                if (node != null)
                {
                    return node.Text;
                }
            }
            catch { }
            return null;
        }

        private void SaveMarkdownToDocument(Word.Document doc, string markdown)
        {
            if (doc == null) return;
            try
            {
                var existing = FindExistingPart(doc);
                if (existing != null)
                {
                    existing.Delete();
                }
                var xml = BuildMarkdownXml(markdown ?? string.Empty);
                doc.CustomXMLParts.Add(xml);
            }
            catch { }
        }

        private Office.CustomXMLPart FindExistingPart(Word.Document doc)
        {
            try
            {
                Office.CustomXMLParts parts = doc.CustomXMLParts;
                foreach (Office.CustomXMLPart p in parts)
                {
                    try
                    {
                        var root = p.DocumentElement;
                        if (root != null && string.Equals(root.NamespaceURI, Services.DocumentSyncService.NamespaceUri, StringComparison.OrdinalIgnoreCase))
                        {
                            return p;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private string BuildMarkdownXml(string content)
        {
            return "<md:markdown xmlns:md='" + Services.DocumentSyncService.NamespaceUri + "'>" +
                "<md:content><![CDATA[" + content + "]]></md:content>" +
                "</md:markdown>";
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
