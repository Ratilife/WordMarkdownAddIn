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

        /// <summary>
        /// Словарь для сопоставления окон Word (<see cref="Word.Window"/>) и связанных с ними
        /// настраиваемых панелей задач (<see cref="Microsoft.Office.Tools.CustomTaskPane"/>).
        /// Используется для управления отдельными панелями Markdown в多个 открытых документах Word.
        /// </summary>
        /// <remarks>
        /// Ключом является объект окна Word, значением - экземпляр панели задач,
        /// созданный для этого окна. Позволяет получить доступ к конкретной панели,
        /// связанной с активным окном, через свойство <see cref="MarkdownPane"/>.
        /// </remarks>
        private static Dictionary<Word.Window, CustomTaskPane> _markdownPanes = new Dictionary<Word.Window, CustomTaskPane>();                      //Управление UI панели (видимость, размер, позиция)

        /// <summary>
        /// Словарь для сопоставления окон Word (<see cref="Word.Window"/>) и связанных с ними
        /// экземпляров пользовательского элемента управления панели задач (<see cref="Controls.TaskPaneControl"/>).
        /// </summary>
        /// <remarks>
        /// Ключом является объект окна Word, значением - экземпляр <c>TaskPaneControl</c>,
        /// созданный для этого окна. Используется для управления отдельными экземплярами
        /// редактора Markdown в открытых документах Word, обеспечивая независимую
        /// работу с Markdown в каждом активном окне.
        /// </remarks>
        private static Dictionary<Word.Window, Controls.TaskPaneControl> _paneControls = new Dictionary<Word.Window, Controls.TaskPaneControl>();   // Работа с содержимым (получение/установка Markdown, вставка элементов) 

        // Статические свойства для обратной совместимости - возвращают панель для активного документа

        /// <summary>
        /// Возвращает экземпляр настраиваемой панели задач (CustomTaskPane), 
        /// связанной с текущим активным окном Microsoft Word.
        /// </summary>
        /// <remarks>
        /// Использует словарь _markdownPanes для получения панели, 
        /// ассоциированной с ActiveWindow. 
        /// Возвращает null, если активное окно отсутствует или панель не была создана.
        /// </remarks>
        /// <returns>
        /// Объект <see cref="Microsoft.Office.Tools.CustomTaskPane"/>, 
        /// представляющий панель задач Markdown для активного окна, 
        /// или <c>null</c>, если панель недоступна.
        /// </returns>
        public static CustomTaskPane MarkdownPane                                                       
        { 
            get 
            {
                // Проверяет, существует ли активное окно в приложении Word.
                if (Instance?.Application?.ActiveWindow != null)           
                {
                    // Пытается получить панель задач из словаря по ключу активного окна.
                    _markdownPanes.TryGetValue(Instance.Application.ActiveWindow, out var pane);  
                    return pane;       // Возвращает найденную панель или null, если ключ отсутствует.
                }
                return null;  // Возвращает null, если Instance, Application или ActiveWindow равны null.
            } 
        }


        /// <summary>
        /// Возвращает экземпляр пользовательского элемента управления панели задач (<see cref="Controls.TaskPaneControl"/>),
        /// связанного с текущим активным окном Microsoft Word.
        /// </summary>
        /// <remarks>
        /// Использует словарь <c>_paneControls</c> для получения элемента управления,
        /// ассоциированного с <c>ActiveWindow</c>.
        /// Возвращает <c>null</c>, если активное окно отсутствует или элемент управления
        /// для этого окна не был создан.
        /// </remarks>
        /// <returns>
        /// Объект <see cref="Controls.TaskPaneControl"/>, представляющий элемент управления
        /// редактора Markdown для активного окна, или <c>null</c>, если элемент управления недоступен.
        /// </returns>
        public static Controls.TaskPaneControl PaneControl 
        { 
            get 
            {
                // Проверяет, существует ли активное окно в приложении Word.
                if (Instance?.Application?.ActiveWindow != null)   
                {
                    // Пытается получить элемент управления из словаря по ключу активного окна.
                    _paneControls.TryGetValue(Instance.Application.ActiveWindow, out var control);   
                    return control; // Возвращает найденный элемент управления или null, если ключ отсутствует.
                }
                return null;  // Возвращает null, если Instance, Application или ActiveWindow равны null.
            } 
        }

        /// <summary>
        /// Предоставляет доступ к экземпляру настраиваемой ленты (<see cref="MarkdownRibbon"/>) надстройки.
        /// </summary>
        /// <value>
        /// Объект <see cref="MarkdownRibbon"/>, представляющий пользовательский интерфейс ленты Word,
        /// связанный с этой надстройкой. Устанавливается только внутри класса <see cref="ThisAddIn"/>.
        /// </value>
        public static MarkdownRibbon Ribbon { get; private set; }

        /// <summary>
        /// Словарь для хранения пользовательских свойств и настроек надстройки, связанных с активным сеансом.
        /// </summary>
        /// <value>
        /// Объект <see cref="Dictionary{TKey, TValue}"/> с ключами типа <see cref="string"/>
        /// и значениями типа <see cref="object"/>, который можно использовать для сохранения
        /// произвольных данных (например, настроек пользовательского интерфейса, состояния).
        /// Устанавливается и изменяется только внутри класса <see cref="ThisAddIn"/>.
        /// </value>
        public Dictionary<string, object> Properties { get; private set; }          //Словарь хранит настройки панели, которые пользователь изменяет:Ширина панели (PaneWidth); Видимость панели (PaneVisible)

        /// <summary>
        /// Предоставляет глобальный доступ к текущему экземпляру надстройки <see cref="ThisAddIn"/>.
        /// </summary>
        /// <value>
        /// Объект <see cref="ThisAddIn"/>, представляющий запущенный экземпляр надстройки.
        /// Используется другими модулями для взаимодействия с основным классом надстройки.
        /// Устанавливается только внутри класса <see cref="ThisAddIn"/>.
        /// </value>
        public static ThisAddIn Instance { get; private set; }

        /// <summary>
        /// Обработчик события <c>Startup</c> для надстройки <c>ThisAddIn</c>.
        /// Выполняет начальную инициализацию компонентов надстройки при её запуске.
        /// </summary>
        /// <remarks>
        /// Метод выполняет следующие шаги:
        /// 1. Инициализирует словарь <c>Properties</c> для хранения настроек.
        /// 2. Устанавливает статическое свойство <c>Instance</c> для глобального доступа к экземпляру надстройки.
        /// 3. Подписывается на события приложения Word (<c>DocumentOpen</c>, <c>WindowActivate</c>, <c…   для создания или получения панели задач, связанной с этим окном.
        /// 5. Создает и инициализирует экземпляр пользовательской ленты <c>MarkdownRibbon</c>.
        /// В каждом блоке инициализации используется базовая обработка исключений <c>try-catch</c>,
        /// которая в текущей реализации подавляет ошибки.
        /// </remarks>
        /// <param name="sender">Объект, инициировавший событие (обычно сам экземпляр <c>ThisAddIn</c>).</param>
        /// <param name="e">Аргументы события (не используются в данном методе).</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 1. Инициализируем словарь свойств
            Properties = new Dictionary<string, object>();

            // 2. Сохраняем экземпляр для глобального доступа
            Instance = this;

            // 3. Подписываемся на события Word для управления панелями
            try
            {
                this.Application.DocumentOpen += Application_DocumentOpen;                  //Срабатывает при открытии документа.
                this.Application.WindowActivate += Application_WindowActivate;              //Срабатывает при активации окна Word.
                this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;    //Срабатывает перед закрытием документа.
                this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;      //Срабатывает перед сохранением документа (используется, например, для автоматического сохранения Markdown).
            }
            catch { }

            // 4. Создаем панель для активного документа (если он есть)
            try
            {
                if (this.Application.ActiveWindow != null)
                {
                    EnsurePaneForWindow(this.Application.ActiveWindow);     //Этот метод отвечает за создание и настройку панели задач (TaskPane) с редактором Markdown для текущего активного окна.
                }
            }
            catch { /* Игнорируем ошибки при старте */ }

            // 5. Инициализируем ленту Ribbon
            try
            {
                Ribbon = new MarkdownRibbon();   // Создается новый экземпляр пользовательской ленты MarkdownRibbon и сохраняется в поле Ribbon. Это добавляет на ленту Word новые вкладки, группы и кнопки, определенные в этой ленте.
            }
            catch { }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 1. Отписываемся от событий Word. Это предотвращает вызов соответствующих обработчиков после выгрузки надстройки.
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
                        this.Properties["PaneWidth"] = firstPane.Width;     // Сохранили ширину
                        this.Properties["PaneVisible"] = firstPane.Visible; // Сохранили видимость
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
