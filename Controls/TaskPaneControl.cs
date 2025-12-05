using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using WordMarkdownAddIn.Properties;

namespace WordMarkdownAddIn.Controls
{
    public class TaskPaneControl: UserControl
    {
        private readonly WebView2 _webView;
        private readonly Services.MarkdownRenderService _renderer;
        private string _latestMarkdown = string.Empty;  // Инициализация пустой строкой  Локальный кэш для быстрого доступа
        private bool _coreReady = false;                //Сигнализирует, что WebView2 полностью инициализирован; 

        public TaskPaneControl() 
        {
            _renderer = new Services.MarkdownRenderService();
            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_webView);                     // Добавляет WebView2 на UserControl   Controls - это коллекция всех дочерних элементов управления
            Load += OnLoadAsync;                        // Подписываем метод OnLoadAsync на событие Load
        }

        private async void OnLoadAsync(object sender, EventArgs e) 
        {
            //выполняет  асинхронную инициализацию WebView2 и настройку его параметров
            await _webView.EnsureCoreWebView2Async();                                                   //Асинхронно инициализирует движок WebView2
                                                                                                        //Загружает WebView2 Runtime (если не установлен)
                                                                                                        //Создает браузерный процесс и окружение
                                                                                                        //await означает "ждать завершения без блокировки UI"
                                                                                                        //Без этой строки _webView.CoreWebView2 будет null          
            
            _coreReady = true;                                                                          //Разрешает выполнение методов, которые работают с WebView2
            _webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;                //Подписывается на событие получения сообщений из JavaScript    
            _webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            _webView.CoreWebView2.NavigateToString(BuildHtmlShell());
        }


        /// <summary>
        /// Обрабатывает сообщения, полученные от Markdown-редактора (JavaScript) и соответствующим образом обрабатывает их в C#
        /// 
        /// Полный процесс передачи данных JavaScript → C#:
        /// 1. JavaScript отправляет: "тип|данныеВBase64" (например: "mdChanged|SGVsbG8=")
        /// 2. → C# разделяет по '|': ["mdChanged", "SGVsbG8="]
        /// 3. → Декодирование Base64: "SGVsbG8=" → "Hello"
        /// 4. → Обработка в зависимости от типа сообщения
        /// 
        /// Полный процесс передачи данных C# → JavaScript:
        /// 1. HTML строка: "<p>Hello</p>"
        /// 2. → UTF8 байты: [60, 112, 62, 72, 101, 108, 108, 111, 60, 47, 112, 62]
        /// 3. → Base64: "PHA+SGVsbG88L3A+"
        /// 4. → JS вызов: window.renderHtml(atob('PHA+SGVsbG88L3A+'))
        /// 5. → JS декодирует: atob('PHA+SGVsbG88L3A+') → "<p>Hello</p>"
        /// 6. → Отображается в preview
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Аргументы события, содержащие полученное сообщение</param>
        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e) 
        {
            //Метод получает сообщения от Markdown-редактора (JavaScript) и соответствующим образом обрабатывает их в C#.
            try 
            { 
                var json = e.TryGetWebMessageAsString();                                                //  извлекает текст сообщения из JS формат: "тип|данныеВBase64"
                if (string.IsNullOrEmpty(json)) return;                                                 // Если сообщение пустое или null - выходим из метода
                var parts = json.Split(new[] { '|' }, 2);                                               // Split - разделяет строку по символу | на 2 части
                                                                                                        // Пример: "mdChanged|SGVsbG8=" → ["mdChanged", "SGVsbG8="]
                if (parts.Length != 2) return;                                                          // Если сообщение не соответствует формату - игнорируем
                var type = parts[0];                                                                    // тип сообщения (например: "mdChanged")
                var payload = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1]));              // декодированные данные из Base64
                                                                                                        // Convert.FromBase64String() - преобразует Base64 в байты
                                                                                                        // Encoding.UTF8.GetString() - преобразует байты в строку UTF-8
                //Обработка изменения Markdown
                if (type == "mdChanged")                                                                // Проверка типа - если это сообщение об изменении Markdown
                {
                    _latestMarkdown = payload;                                                          // Сохранение - обновление кэша
                    var html = _renderer.RenderoHtml(payload);                                          // Конвертация - Markdown → HTML
                    PostRenderHtml(html);                                                               // Отправка - показ HTML в preview
                }
                
                // Обработка изменения режима отображения
                if (type == "viewModeChanged")
                {
                    // Сохраняем режим в настройках приложения
                    try
                    {
                        Settings.Default.ViewMode = payload;
                        Settings.Default.Save();
                    }
                    catch { /* Игнорируем ошибки сохранения настроек */ }
                }

            }
            catch {/* ignore malformed messages */ }
        }

        private void PostRenderHtml(string html) 
        {
            if (!_coreReady || _webView == null) return;                                                // Проверка готовности WebView2
            try
            {
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return;
                
                var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(html));                         // Кодирование HTML в Base64
                                                                                                        // Encoding.UTF8.GetBytes(html) - преобразует строку HTML в байты UTF-8
                                                                                                        // Convert.ToBase64String() - кодирует байты в строку Base64
                                                                                                        // Пример: <p>Hello</p> → "PHA+SGVsbG88L3A+"
                coreWebView2.ExecuteScriptAsync($"window.renderHtml(atob('{b64}'));void(0);");          // Выполнение JavaScript кода
                                                                                                        // ExecuteScriptAsync() - асинхронно выполняет JavaScript код в WebView2
                                                                                                        // $"..." - строковая интерполяция C# для подстановки переменных
                                                                                                        // window.renderHtml() - вызов JavaScript функции из C#
                                                                                                        // atob('{b64}') - JavaScript функция декодирования Base64
                                                                                                        // void(0); - предотвращает возврат значения (оптимизация)
            }
            catch (InvalidCastException)
            {
                // WebView2 еще не готов, игнорируем
                return;
            }
            catch (Exception)
            {
                // Другие ошибки тоже игнорируем
                return;
            }
        }

        /// <summary>
        /// Устанавливает Markdown-контент в редакторе веб-интерфейса.
        /// Преобразует входную строку в Base64 для безопасной передачи в JavaScript,
        /// затем выполняет скрипт для установки значения в редакторе.
        /// 
        /// Процесс преобразования:
        /// - Исходный текст: "Hello **World**"
        /// - Байты: [72, 101, 108, 108, 111, 32, 42, 42, 87, 111, 114, 108, 100, 42, 42]
        /// - Base64: "SGVsbG8gKipXb3JsZCoq"
        /// - JS вызов: window.editorSetValue(atob('SGVsbG8gKipXb3JsZCoq'))
        /// - JS: atob('SGVsbG8gKipXb3JsZCoq') → "Hello **World**"
        /// - Редактор: получает декодированный текст
        /// </summary>
        /// <param name="markdown">Markdown-текст для установки в редактор. Если null, используется пустая строка.</param>
        public void SetMarkdown(string markdown) 
        {
            //Метод обновляет содержимое Markdown-редактора в WebView2 и сохраняет копию в памяти C#.
            _latestMarkdown = markdown ?? string.Empty;                                                                 // Сохранение в кэш C#
                                                                                                                        // markdown ?? string.Empty - если null, использует пустую строку            
                                                                                                                        // _latestMarkdown - локальная переменная для хранения текущего текста
                                                                                                                        // Цель: Быстрый доступ к тексту без запроса к JavaScript
            if (!_coreReady || _webView == null) return;                                                               // Проверка готовности WebView2
            try
            {
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return;
                
                var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(_latestMarkdown));                              // Кодирование в Base64
                                                                                                                        // Encoding.UTF8.GetBytes() - преобразует строку в байты UTF-8
                                                                                                                        // Convert.ToBase64String() - кодирует байты в Base64 строку
                                                                                                                        // Пример: "Hello" → "SGVsbG8="    
                coreWebView2.ExecuteScriptAsync($"window.editorSetValue(atob('{b64}'));void(0);");                      // Отправка в JavaScript
                                                                                                                        // ExecuteScriptAsync() - выполняет JavaScript код асинхронно
                                                                                                                        // window.editorSetValue() - JS функция для установки значения редактора
                                                                                                                        // atob('{b64}') - JS функция декодирования Base64
                                                                                                                        // void(0); - предотвращает возврат значения (оптимизация)
            }
            catch (InvalidCastException)
            {
                // WebView2 еще не готов, игнорируем
                return;
            }
            catch (Exception)
            {
                // Другие ошибки тоже игнорируем
                return;
            }
        }

        public string GetCachedMarkdown() => _latestMarkdown;          //предоставляет мгновенный доступ к последнему известному состоянию Markdown-текста без обращения к JavaScript.

        //Метод предоставляет надежный способ получения актуального Markdown-текста,используя кэш, но при необходимости запрашивая данные из JavaScript.
        public async Task<string> GetMarkdownAsync()
        {
            //  Возвращает кэшированное значение (синхронизируется с помощью mdChanged). При необходимости можно вернуться к JS-запросу.
            if (!string.IsNullOrEmpty(_latestMarkdown)) return _latestMarkdown;                         //!string.IsNullOrEmpty() - проверяет что кэш не пустой
                                                                                                        // return _latestMarkdown - мгновенный возврат из кэша
            if (_coreReady && _webView != null)                                                         //_coreReady - флаг инициализации WebView2
            {
                try
                {
                    var coreWebView2 = _webView.CoreWebView2;
                    if (coreWebView2 != null)
                    {
                        var js = await coreWebView2.ExecuteScriptAsync("window.editorGetValue()");
                        return UnquoteJsonString(js);                                                   // Обработка результата JavaScript
                    }
                }
                catch (InvalidCastException)
                {
                    // WebView2 еще не готов
                    return string.Empty;
                }
                catch (Exception)
                {
                    // Другие ошибки
                    return string.Empty;
                }
            }
            return string.Empty;                                                                        // Если WebView2 не готов и кэш пустой
        }

        // Метод обрабатывает строки, возвращаемые из JavaScript, которые могут быть в JSON-формате с экранированными символами.
        private static string UnquoteJsonString(string jsonQuoted)
        {
            if (string.IsNullOrEmpty(jsonQuoted)) return string.Empty;                                                          //Защита от null или пустых входных данны
            var s = jsonQuoted;                                                                                                 //Создает копию для безопасного изменения
            if (s.StartsWith("\"") && s.EndsWith("\"")) s = s.Substring(1, s.Length - 2);                                       // StartsWith(""") - проверяет начинается ли с кавычки
                                                                                                                                // EndsWith(""") - проверяет заканчивается ли кавычкой
                                                                                                                                // Substring(1, s.Length - 2) - удаляет первую и последнюю кавычки
                                                                                                                                // Пример: "Hello" → Hello        
            s = s.Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t").Replace("\\\"", "\"").Replace("\\\\", "\\");   // \\n → \n - новая строка
                                                                                                                                // \\r → \r - возврат каретки
                                                                                                                                // \\t → \t - табуляция
                                                                                                                                // \\" → " - кавычка
                                                                                                                                // \\\\ → \\ - обратный слеш
            return s;                                                                                                           // Возвращает обработанную строку
        }

        //Метод добавляет префикс и суффикс вокруг выделенного текста в редакторе (например, для жирного текста или кода)
        public void InsertInline(string prefix, string suffix)
        {
            if (!_coreReady || _webView == null) return;                                                                        //_coreReady - флаг инициализации WebView2
            try
            {
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return;
                
                var p = Convert.ToBase64String(Encoding.UTF8.GetBytes(prefix ?? string.Empty));                                 //prefix ?? string.Empty - если null, использует пустую строку
                                                                                                                                //Encoding.UTF8.GetBytes() - преобразует строку в байты UTF-8
                                                                                                                                //Convert.ToBase64String() - кодирует байты в Base64
                                                                                                                                // Пример: "**" → "Kg=="
                var s = Convert.ToBase64String(Encoding.UTF8.GetBytes(suffix ?? string.Empty));                                 // Аналогично префиксу, но для суффикса
                                                                                                                                // Пример: "**" → "Kg=="
                coreWebView2.ExecuteScriptAsync($"window.insertAroundSelection(atob('{p}'), atob('{s}'));void(0);");             // Отправка в JavaScript
                                                                                                                                // ExecuteScriptAsync() - выполняет JS код асинхронно
                                                                                                                                // window.insertAroundSelection() - JS функция для вставки вокруг выделения
                                                                                                                                // atob('{p}') - JS декодирование Base64 префикса
                                                                                                                                // atob('{s}') - JS декодирование Base64 суффикса
                                                                                                                                // void(0); - предотвращает возврат значения
            }
            catch (InvalidCastException)
            {
                // WebView2 еще не готов, игнорируем
                return;
            }
            catch (Exception)
            {
                // Другие ошибки тоже игнорируем
                return;
            }
        }
        //Метод вставляет предопределенные блоки Markdown(заголовки, списки, таблицы) в текущую позицию курсора.
        public void InsertSnippet(string snippet)
        {
            if (!_coreReady || _webView == null) return;                                                                         // _coreReady - флаг инициализации WebView2    
            try
            {
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return;
                
                var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(snippet ?? string.Empty));                              // Кодирование сниппета в Base64
                                                                                                                                // snippet ?? string.Empty - если null, использует пустую строку
                                                                                                                                // Encoding.UTF8.GetBytes() - преобразует строку в байты UTF-8
                                                                                                                                // Convert.ToBase64String() - кодирует байты в Base64
                                                                                                                                // Пример: "# Заголовок\n" → "IyDQl9Cw0LPQu9C10LbQvtC8Cg=="
                coreWebView2.ExecuteScriptAsync($"window.insertSnippet(atob('{b64}'));void(0);");                                // Отправка в JavaScript
                                                                                                                                // ExecuteScriptAsync() - выполняет JS код асинхронно
                                                                                                                                // window.insertSnippet() - JS функция для вставки сниппета
                                                                                                                                // atob('{b64}') - JS декодирование Base64
                                                                                                                                // void(0); - предотвращает возврат значения
            }
            catch (InvalidCastException)
            {
                // WebView2 еще не готов, игнорируем
                return;
            }
            catch (Exception)
            {
                // Другие ошибки тоже игнорируем
                return;
            }
        }

        //Метод генерирует и вставляет заголовок Markdown соответствующего уровня (от H1 до H6).
        public void InsertHeading(int level)
        {
            //Гарантирует, что уровень всегда будет между 1 и 6
            if (level < 1) level = 1;                                                                                           // level < 1 - если уровень меньше 1, устанавливается 1        
            if (level > 6) level = 6;                                                                                           // level > 6 - если уровень больше 6, устанавливается 6
            InsertSnippet("\n" + new string('#', level) + " ");                                                                 // new string('#', level) - создает строку из # символов длиной level
                                                                                                                                // level = 1 → "#"; level = 2 → "##"; level = 3 → "###"
                                                                                                                                // "\n" - начинается с новой строки (чтобы не склеилось с предыдущим текстом)
                                                                                                                                // " " - заканчивается пробелом (обязательно в Markdown для заголовков)
        }

        //Метод вставляет Markdown-синтаксис для начала маркированного списка (bullet list).
        public void InsertBulletList()
        {
            InsertSnippet("\n- ");                                                                                              // "\n- " - строка, содержащая Markdown-синтаксис для маркированного списка
                                                                                                                                // "\n" - перевод строки (начать с новой строки)
                                                                                                                                // "-" - дефис (Markdown-синтаксис для элемента списка)
                                                                                                                                // " " - пробел (обязателен после дефиса в Markdown)
        }

        //Метод вставляет Markdown-синтаксис для начала нумерованного списка (numbered list).
        public void InsertNumberedList()
        {
            InsertSnippet("\n1. ");                                                                                             // "\n1. " - строка, содержащая Markdown-синтаксис для нумерованного списка
        }

        //Метод вставляет Markdown-синтаксис для чекбокса (task list item) с возможностью выбора состояния - отмечен или нет.
        public void InsertCheckbox(bool isChecked)
        {
            InsertSnippet(isChecked ? "\n- [x] " : "\n- [ ] ");                                                                 // Условный оператор для выбора состояния
                                                                                                                                // isChecked - булев параметр (true/false)
                                                                                                                                // ? : - тернарный оператор (if-else в одной строке)
                                                                                                                                // Если isChecked = true → возвращает "\n- [x] "
                                                                                                                                // Если isChecked = false → возвращает "\n- [ ] "
        }

        //Метод создает шаблон таблицы в Markdown-формате с заданным количеством строк и столбцов.
        public void InsertTable(int rows, int cols)
        {
            // Валидация параметров; Гарантирует, что таблица будет хотя бы 2x2
            if (rows < 2) rows = 2;                                                                                             // rows < 2 - если строк меньше 2, устанавливает 2 (минимум)    
            if (cols < 2) cols = 2;                                                                                             // cols < 2 - если столбцов меньше 2, устанавливает 2 (минимум)        
            var sb = new StringBuilder();                                                                                       //StringBuilder - эффективный способ построения больших строк
            // Заголовок таблицы
            for (int c = 0; c < cols; c++) sb.Append("| Header").Append(c + 1).Append(' ');                                     // Цикл по столбцам - создает ячейки заголовка
                                                                                                                                // "| Header" - начало ячейки
                                                                                                                                // Append(c + 1) - номер столбца (1-based)
                                                                                                                                // Append(' ') - пробел для читаемости
            sb.AppendLine("|");                                                                                                 // AppendLine("|") - закрывающая вертикальная черта
            for (int c = 0; c < cols; c++) sb.Append("| --- ");                                                                 // "| --- " - Markdown-синтаксис для разделителя столбцов
            sb.AppendLine("|");
            //Тело таблицы
            for (int r = 0; r < rows - 1; r++)                                                                                  // rows - 1 - потому что первая строка уже создана как заголовок
            {                                                                                                                   // Вложенные циклы - строки → столбцы
                for (int c = 0; c < cols; c++) sb.Append("| cell ");                                                            // "| cell " - шаблон ячейки с текстом
                sb.AppendLine("|");                                                                                             // AppendLine("|") - закрытие строки
            }
            InsertSnippet("\n" + sb.ToString() + "\n");                                                                         // Вставка таблицы
        }

        //Метод создает и вставляет Markdown-синтаксис для гиперссылки с указанным текстом и URL.
        public void InsertLink(string text, string url)
        {
            InsertSnippet($"[{text}]({url})");                                                      //$ - строковая интерполяция (C# 6+)
                                                                                                    //[ - открывающая квадратная скобка (начало текста ссылки)
                                                                                                    //{text} - вставляет параметр text (анкор ссылки)
                                                                                                    //] - закрывающая квадратная скобка (конец текста ссылки)
                                                                                                    //( - открывающая круглая скобка (начало URL)
                                                                                                    //{url} - вставляет параметр url (адрес ссылки)
                                                                                                    //) - закрывающая круглая скобка (конец URL)
                                                                                                    // Результат: [Google](https://google.com)

        }

        //Метод создает и вставляет Markdown-синтаксис для отображения изображения с указанным альтернативным текстом и путем к файлу.
        public void InsertImage(string alt, string path)
        {
            InsertSnippet($"![{alt}]({path})");                                                     //! - восклицательный знак (обозначает изображение, а не ссылку)
                                                                                                    //{alt} - вставляет параметр alt (альтернативный текст)
                                                                                                    //{path} - вставляет параметр path (путь или URL изображения)
                                                                                                    // Результат: ![Логотип](/images/logo.png)
        }

        //Метод создает и вставляет Markdown-синтаксис для блока кода с подсветкой синтаксиса для конкретного языка программирования.
        public void InsertCodeBlock(string language)                                                //"\n" - перевод строки перед блоком (чтобы отделить от предыдущего текста)
        {                                                                                           //``` - тройные backticks (начало/конец блока кода в Markdown)
            InsertSnippet($"\n```{language}\n\n```\n");                                             //{language} - вставляет параметр language (язык программирования)
                                                                                                    //"\n\n" - два перевода строки (место для ввода кода)
                                                                                                    //``` - закрывающие тройные backticks
                                                                                                    //"\n" - перевод строки после блока
                                                                                                    // Результат:
                                                                                                    // ```csharp
                                                                                                    // 
                                                                                                    // ```
        }

        //Метод вставляет готовый пример диаграммы на языке Mermaid для демонстрации возможностей визуализации.
        public void InsertMermaidSample()                                                           //"```mermaid\n" - начало блока кода с указанием языка mermaid
        {                                                                                           //"graph TD; A-->B; A-->C; B-->D; C-->D;\n" - код Mermaid-диаграммы
            InsertSnippet("\n```mermaid\ngraph TD; A-->B; A-->C; B-->D; C-->D;\n```\n");            //"```\n" - закрытие блока кода
        }

        public void InsertMermaid(string mermaid_text)                                             
        {                                                                                           
            InsertSnippet($"\n```mermaid\n{mermaid_text}\n```\n");            
        }

        //Метод вставляет демонстрационную математическую формулу для показа возможностей рендеринга математических выражений в Markdown.
        public void InsertMathSample()                                                              //"\n" - перевод строки перед формулой (отделяет от предыдущего текста)            
        {                                                                                           //"$$" - двойной знак доллара (обозначает блочную математическую формулу в LaTeX)
            InsertSnippet("\n$$\\int_{0}^{1} x^2 \\; dx = \\tfrac{1}{3}$$\n");                      //"\int_{0}^{1}" - интеграл от 0 до 1 (∫₀¹)
                                                                                                    //" x^2 " - x в квадрате (x²)  
                                                                                                    //"\;" - пробел среднего размера в LaTeX
                                                                                                    //" dx " - дифференциал dx
                                                                                                    //" = " - знак равенства
                                                                                                    //"\tfrac{1}{3}" - дробь одна треть (½)
                                                                                                    //"$$" - закрывающие двойные знаки доллара
        }                                                                                           //"\n" - перевод строки после формулы

        public void InsertMath(string math_text)                                                                          
        {                                                                                           
            InsertSnippet($"\n$${math_text}$$\n");                      
                                                
        }                                                                                           

        //Метод позволяет пользователю сохранить содержимое Markdown-редактора в файл на диске и одновременно синхронизировать с текущим Word-документом.
        public async void SaveMarkdownFile()                                                        //
        {
            //Получение Markdown-контента
            var md = await GetMarkdownAsync();                                                      // GetMarkdownAsync() - асинхронно получает текущий текст из редактора
                                                                                                    // await - ожидает завершения асинхронной операции
            //Создание диалога сохранения файла                                                     // md - содержит актуальный Markdown-текст
            using (var dlg = new SaveFileDialog())                                                  //SaveFileDialog() - стандартный диалог сохранения файла Windows
            {                                                                                       //using - гарантирует корректное освобождение ресурсов
                dlg.Filter = "Markdown (*.md)|*.md|All files (*.*)|*.*";                            //Filter - определяет типы файлов в диалоге;"Markdown (.md)|.md" - показывает/фильтрует .md файлы;"All files (.)|." - опция показа всех файлов
                dlg.FileName = "document.md";                                                       //FileName - предлагаемое имя файла; "document.md" - имя по умолчанию
                if (dlg.ShowDialog() == DialogResult.OK)                                            //ShowDialog() - показывает модальный диалог;
                {                                                                                   // DialogResult.OK - пользователь выбрал файл и нажал "Сохранить"
                    File.WriteAllText(dlg.FileName, md ?? string.Empty, new UTF8Encoding(false));   //File.WriteAllText() - записывает весь текст в файл
                                                                                                    //dlg.FileName - путь к выбранному файлу
                                                                                                    //md ?? string.Empty - защита от null
                                                                                                    //new UTF8Encoding(false) - кодировка UTF-8 без BOM
                    Services.DocumentSyncService.SaveMarkdownToActiveDocument(                      //SaveMarkdownToActiveDocument() - мой  сервис синхронизации
                        Globals.ThisAddIn.Application, md ?? string.Empty);                         //Globals.ThisAddIn.Application - ссылка на приложение Word
                }                                                                                   //md ?? string.Empty - тот же Markdown-контент
            }
        }

        public async void OpenMarkdownFile()
        {
            //Метод позволяет пользователю выбрать файл Markdown (.md), загрузить его содержимое в редактор и синхронизировать с текущим Word документом.
            using (var dlg = new OpenFileDialog())                                                                      // Создание диалога открытия файла
            {                                                                                                           // OpenFileDialog() - стандартный диалог выбора файла Windows; using - гарантирует корректное освобождение ресурсов после использования
                dlg.Filter = "Markdown (*.md)|*.md|All files (*.*)|*.*";                                                // Настройка фильтра файлов; определяет типы файлов в диалоге; Формат: "Описание|шаблон|описание|шаблон..."
                if (dlg.ShowDialog() == DialogResult.OK)                                                                // ShowDialog() - показывает модальный диалог выбора файла; DialogResult.OK - пользователь выбрал файл и нажал "Открыть"
                {
                    var text = File.ReadAllText(dlg.FileName, new UTF8Encoding(false));                                 // Чтение содержимого файла
                                                                                                                        // dlg.FileName - путь к выбранному файлу
                                                                                                                        // File.ReadAllText() - читает весь текст файла
                                                                                                                        // new UTF8Encoding(false) - указывает кодировку UTF-8 без BOM
                                                                                                                        // BOM (Byte Order Mark) - не нужен для чистого UTF-8
                    SetMarkdown(text);                                                                                  // Загрузка текста в редактор
                    Services.DocumentSyncService.SaveMarkdownToActiveDocument(Globals.ThisAddIn.Application, text);     // Синхронизация с Word документом
                                                                                                                        // Services.DocumentSyncService - сервис синхронизации
                                                                                                                        // SaveMarkdownToActiveDocument() - сохраняет Markdown в текущий Word документ
                                                                                                                        // Globals.ThisAddIn.Application - ссылка на приложение Word
                                                                                                                        // text - содержимое Markdown файла
                }
            }
        }

        /// <summary>
        /// Устанавливает режим отображения панели (Split, Markdown-only, HTML-only)
        /// </summary>
        /// <param name="mode">Режим отображения: "split", "markdown" или "html"</param>
        public void SetViewMode(string mode)
        {
            // Безопасная проверка готовности WebView2
            if (!_coreReady || _webView == null) return;
            
            try
            {
                // Проверяем наличие CoreWebView2 безопасным способом
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return;
                
                // Валидация режима
                if (mode != "split" && mode != "markdown" && mode != "html")
                    mode = "split";
                
                // Вызов JavaScript функции для переключения режима
                coreWebView2.ExecuteScriptAsync($"window.setViewMode('{mode}');void(0);");
            }
            catch (InvalidCastException)
            {
                // WebView2 еще не готов, игнорируем ошибку
                return;
            }
            catch (Exception)
            {
                // Другие ошибки тоже игнорируем
                return;
            }
        }

        /// <summary>
        /// Получает текущий режим отображения панели
        /// </summary>
        /// <returns>Текущий режим: "split", "markdown" или "html"</returns>
        public async Task<string> GetCurrentViewModeAsync()
        {
            if (!_coreReady || _webView == null) return "split";
            
            try
            {
                // Безопасная проверка CoreWebView2
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return "split";
                
                var result = await coreWebView2.ExecuteScriptAsync("window.getViewMode()");
                var mode = UnquoteJsonString(result);
                
                // Валидация результата
                if (mode == "split" || mode == "markdown" || mode == "html")
                    return mode;
                
                return "split";
            }
            catch (InvalidCastException)
            {
                // WebView2 еще не готов
                return "split";
            }
            catch
            {
                return "split";
            }
        }

        private string BuildHtmlShell()
        {
            return @"<!DOCTYPE html>
        <html>
        <head>
            <meta charset=""utf-8""/>
            <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
            <meta name=""viewport"" content=""width=device-width, initial-scale=1"" />
            <title>Markdown Editor</title>
            <style>
                html, body { 
                    height: 100%; 
                    margin: 0; 
                    font-family: Segoe UI, Arial, sans-serif; 
                }   
                
                /* Панель управления режимами */
                .view-controls {
                    display: flex;
                    gap: 4px;
                    padding: 8px;
                    background: #f5f5f5;
                    border-bottom: 1px solid #ddd;
                }
                
                .view-btn {
                    padding: 6px 12px;
                    border: 1px solid #ccc;
                    background: white;
                    cursor: pointer;
                    border-radius: 4px;
                    font-size: 12px;
                    transition: all 0.2s ease;
                }
                
                .view-btn:hover {
                    background: #e0e0e0;
                }
                
                .view-btn.active {
                    background: #0078d4;
                    color: white;
                    border-color: #0078d4;
                }
                
                .container { 
                    display: flex; 
                    height: calc(100% - 45px); 
                }
                
                #editor { 
                    width: 50%; 
                    height: 100%; 
                    border: none; 
                    padding: 12px; 
                    font-family: Consolas, monospace; 
                    font-size: 13px; 
                    box-sizing: border-box; 
                    outline: none; 
                    resize: none; 
                    border-right: 1px solid #ddd; 
                    transition: width 0.3s ease, opacity 0.2s ease;
                }
                
                #preview { 
                    width: 50%; 
                    height: 100%; 
                    overflow: auto; 
                    padding: 16px; 
                    box-sizing: border-box; 
                    transition: width 0.3s ease, opacity 0.2s ease;
                }
                
                /* Режим Split (по умолчанию) */
                .view-mode-split .container {
                    display: flex;
                }
                
                .view-mode-split #editor {
                    width: 50%;
                    display: block;
                }
                
                .view-mode-split #preview {
                    width: 50%;
                    display: block;
                }
                
                /* Режим только Markdown */
                .view-mode-markdown .container {
                    display: flex;
                }
                
                .view-mode-markdown #editor {
                    width: 100%;
                    display: block;
                }
                
                .view-mode-markdown #preview {
                    display: none;
                }
                
                /* Режим только HTML */
                .view-mode-html .container {
                    display: flex;
                }
                
                .view-mode-html #editor {
                    display: none;
                }
                
                .view-mode-html #preview {
                    width: 100%;
                    display: block;
                }
                
                pre { 
                    background: #f6f8fa; 
                    padding: 10px; 
                    overflow: auto; 
                }
                code { 
                    font-family: Consolas, monospace; 
                }
            </style>
        </head>
        <body>
            <div class=""view-controls"">
                <button id=""btn-split"" class=""view-btn active"">Split</button>
                <button id=""btn-markdown"" class=""view-btn"">Markdown</button>
                <button id=""btn-html"" class=""view-btn"">HTML</button>
            </div>
            <div class=""container"">
                <textarea id=""editor"" placeholder=""Введите Markdown...""></textarea>
                <div id=""preview""></div>
            </div>

            <!-- Скрипты -->
            <script src=""https://cdn.jsdelivr.net/npm/dompurify@3.1.0/dist/purify.min.js""></script>
            <script src=""https://cdn.jsdelivr.net/npm/prismjs@1.29.0/prism.min.js""></script>
            <script src=""https://cdn.jsdelivr.net/npm/prismjs@1.29.0/plugins/autoloader/prism-autoloader.min.js""></script>
            <script>Prism.plugins.autoloader.languages_path = 'https://cdn.jsdelivr.net/npm/prismjs@1.29.0/components/';</script>
            <script src=""https://cdn.jsdelivr.net/npm/mermaid@10.9.0/dist/mermaid.min.js""></script>
            <script>mermaid.initialize({ startOnLoad: false, securityLevel: 'strict' });</script>
            <script>window.MathJax = { tex: { inlineMath: [['$', '$'], ['\\\(', '\\\)']] }, svg: { fontCache: 'global' } };</script>
            <script src=""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js""></script>
            
            <script>
                // Базовые переменные
                const editor = document.getElementById('editor');
                const preview = document.getElementById('preview');
                
                // Переменная для текущего режима отображения
                let currentViewMode = 'split'; // по умолчанию
                
                // Флаг для предотвращения рекурсии
                let setViewModeInProgress = false;
                let setViewModeRetryCount = 0;
                const MAX_RETRIES = 5;
                
                // Внутренняя функция переключения режима (без вызова postToHost)
                function applyViewMode(mode, notifyCSharp) {
                    notifyCSharp = notifyCSharp !== false; // по умолчанию true
                    
                    // Валидация режима
                    if (mode !== 'split' && mode !== 'markdown' && mode !== 'html') {
                        console.error('Неверный режим:', mode);
                        mode = 'split';
                    }
                    
                    currentViewMode = mode;
                    const className = 'view-mode-' + mode;
                    
                    try {
                        // Удаляем все классы режимов с body
                        var body = document.body;
                        var currentClass = body.className || '';
                        
                        // Удаляем старые классы режимов
                        currentClass = currentClass.replace(/view-mode-\w+/g, '').trim();
                        
                        // Добавляем новый класс
                        if (currentClass) {
                            body.className = currentClass + ' ' + className;
                        } else {
                            body.className = className;
                        }
                        
                        // Обновляем активную кнопку
                        document.querySelectorAll('.view-btn').forEach(function(btn) {
                            if (btn.classList) {
                                btn.classList.remove('active');
                            } else {
                                btn.className = btn.className.replace(/\s*active\s*/g, ' ').trim();
                            }
                        });
                        
                        const activeBtn = document.getElementById('btn-' + mode);
                        if (activeBtn) {
                            if (activeBtn.classList) {
                                activeBtn.classList.add('active');
                            } else {
                                activeBtn.className = (activeBtn.className + ' active').trim();
                            }
                        }
                        
                        // Сохраняем в localStorage (с обработкой ошибок)
                        try {
                            if (typeof localStorage !== 'undefined' && localStorage !== null) {
                                localStorage.setItem('viewMode', mode);
                            }
                        } catch(e) {
                            // localStorage недоступен в WebView2 для about:blank - это нормально
                        }
                        
                        // Уведомляем C# о изменении режима (только если нужно)
                        if (notifyCSharp) {
                            postToHost('viewModeChanged', mode);
                        }
                        
                        // Отладочный вывод
                        console.log('Режим переключен на:', mode);
                        console.log('Класс body:', document.body.className);
                    } catch(e) {
                        console.error('Ошибка в applyViewMode:', e);
                        // Пытаемся установить класс напрямую
                        try {
                            document.body.className = 'view-mode-' + mode;
                            console.log('Класс установлен напрямую:', document.body.className);
                        } catch(e2) {
                            console.error('Критическая ошибка установки класса:', e2);
                        }
                    }
                }
                
                // Основная функция переключения режима отображения
                function setViewMode(mode) {
                    // Защита от рекурсии
                    if (setViewModeInProgress) {
                        console.warn('setViewMode уже выполняется, пропускаем вызов');
                        return;
                    }
                    
                    // Проверяем, что body существует
                    if (!document || !document.body) {
                        if (setViewModeRetryCount < MAX_RETRIES) {
                            setViewModeRetryCount++;
                            console.log('document.body не найден! Попытка ' + setViewModeRetryCount + ' через 200мс...');
                            setTimeout(function() {
                                setViewMode(mode);
                            }, 200);
                        } else {
                            console.error('Не удалось найти document.body после ' + MAX_RETRIES + ' попыток');
                            setViewModeRetryCount = 0;
                        }
                        return;
                    }
                    
                    // Сбрасываем счетчик при успехе
                    setViewModeRetryCount = 0;
                    setViewModeInProgress = true;
                    
                    try {
                        // Применяем режим с уведомлением C#
                        applyViewMode(mode, true);
                    } finally {
                        // Снимаем флаг после завершения
                        setViewModeInProgress = false;
                    }
                }
                
                // Обработчики кнопок переключения режимов

             
                // Функция инициализации обработчиков кнопок
                function initViewModeButtons() {
                    // Проверяем, что кнопки существуют, перед добавлением обработчиков
                    const btnSplit = document.getElementById('btn-split');
                    const btnMarkdown = document.getElementById('btn-markdown');
                    const btnHtml = document.getElementById('btn-html');
    
                    // Добавляем обработчик только если кнопка найдена
                    if (btnSplit) {
                        btnSplit.addEventListener('click', function() {
                            setViewMode('split');
                        });
                    } else {
                        console.error('Кнопка btn-split не найдена!');
                    }
    
                    if (btnMarkdown) {
                        btnMarkdown.addEventListener('click', function() {
                            setViewMode('markdown');
                        });
                    } else {
                        console.error('Кнопка btn-markdown не найдена!');
                    }
    
                    if (btnHtml) {
                        btnHtml.addEventListener('click', function() {
                            setViewMode('html');
                        });
                    } else {
                        console.error('Кнопка btn-html не найдена!');
                    }
                }

                // Вызываем функцию инициализации сразу
                initViewModeButtons();

                // Функции для вызова из C#
                window.setViewMode = function(mode) {
                    // Защита от рекурсии
                    if (setViewModeInProgress) {
                        console.warn('window.setViewMode: уже выполняется, пропускаем');
                        return;
                    }
                    
                    if (mode === 'split' || mode === 'markdown' || mode === 'html') {
                        // Проверяем, что body существует
                        if (!document || !document.body) {
                            console.error('window.setViewMode: document.body не найден');
                            return;
                        }
                        
                        // Применяем режим БЕЗ уведомления C# (чтобы избежать цикла)
                        setViewModeInProgress = true;
                        try {
                            applyViewMode(mode, false); // false = не уведомлять C#
                        } finally {
                            setViewModeInProgress = false;
                        }
                    }
                };
                
                window.getViewMode = function() {
                    return currentViewMode;
                };

                // Функция для отправки сообщений в C#
                function postToHost(type, text) {
                    try {
                        // Правильное кодирование base64
                        const b64 = btoa(encodeURIComponent(text || ''));
                        if (window.chrome && window.chrome.webview) {
                            window.chrome.webview.postMessage(type + '|' + b64);
                        }
                        else if (window.external && typeof window.external.notify === 'function') {
                            window.external.notify(type + '|' + b64);
                        }
                    } catch(e) { 
                        console.error('Ошибка отправки:', e); 
                    }
                }

                // Уведомление об изменениях с задержкой
                function debounce(fn, ms) { 
                    let t; 
                    return function() { 
                        clearTimeout(t); 
                        t = setTimeout(() => fn.apply(this, arguments), ms); 
                    } 
                }

                function notifyChange() { 
                    postToHost('mdChanged', editor.value); 
                }

                // Слушаем изменения в редакторе
                editor.addEventListener('input', debounce(notifyChange, 120));

                // Методы для вызова из C#
                window.editorSetValue = function(text) { 
                    editor.value = text || ''; 
                    notifyChange(); 
                }
                
                window.editorGetValue = function() { 
                    return editor.value || ''; 
                }

                window.insertAroundSelection = function(prefix, suffix) {
                    prefix = prefix || ''; 
                    suffix = suffix || '';
                    const start = editor.selectionStart || 0;
                    const end = editor.selectionEnd || 0;
                    const val = editor.value;
    
                    editor.value = val.substring(0, start) + 
                        prefix + 
                        val.substring(start, end) + 
                        suffix + 
                        val.substring(end);
    
                    const newPos = start + prefix.length + (end - start);
                    editor.setSelectionRange(newPos, newPos);
                    editor.focus();
                    notifyChange();
                }
                
                window.insertSnippet = function(snippet) {
                    const pos = editor.selectionStart || 0;
                    const val = editor.value;
                    editor.value = val.substring(0, pos) + snippet + val.substring(pos);
                    const newPos = pos + snippet.length;
                    editor.setSelectionRange(newPos, newPos);
                    editor.focus();
                    notifyChange();
                }

                window.renderHtml = function(html) {
                    try {
                        // Базовая очистка и отображение
                        const clean = DOMPurify.sanitize(html || '', { 
                            ADD_ATTR: ['target', 'rel', 'class', 'style', 'id'] 
                        });
                        preview.innerHTML = clean;
                        
                        // Преобразовать блоки кода mermaid в divs
                        const mermaidBlocks = preview.querySelectorAll('pre code.language-mermaid');
                        mermaidBlocks.forEach(code => {
                            const pre = code.parentElement;
                            const wrapper = document.createElement('div');
                            wrapper.className = 'mermaid';
                            wrapper.textContent = code.textContent;
                            pre.replaceWith(wrapper);
                        });
                        
                        Prism.highlightAllUnder(preview);
                        
                        if (window.mermaid) {
                            mermaid.init(undefined, preview.querySelectorAll('.mermaid'));
                        }
                        
                        if (window.MathJax && MathJax.typesetPromise) {
                            MathJax.typesetPromise([preview]).catch(err => console.error(err));
                        }
                    } catch(e) { 
                        console.error('Ошибка рендеринга:', e); 
                    }
                }

                // Инициализация после загрузки
                function initializeApp() {
                    // Проверяем готовность DOM
                    if (!document.body) {
                        console.log('DOM еще не готов, повтор через 200мс...');
                        setTimeout(initializeApp, 200);
                        return;
                    }
                    
                    // Загрузка сохраненного режима из localStorage (с обработкой ошибок)
                    let savedMode = 'split';
                    try {
                        if (typeof localStorage !== 'undefined' && localStorage !== null) {
                            savedMode = localStorage.getItem('viewMode') || 'split';
                            console.log('Загружаем сохраненный режим:', savedMode);
                        } else {
                            console.log('localStorage недоступен, используем режим по умолчанию: split');
                        }
                    } catch(e) {
                        // localStorage недоступен в WebView2 для about:blank - это нормально
                        console.log('localStorage недоступен (это нормально для WebView2), используем split');
                        savedMode = 'split';
                    }
                    
                    // Устанавливаем режим
                    setViewMode(savedMode);

                    // Инициализируем кнопки еще раз (на всякий случай)
                    initViewModeButtons();
                    
                    // Финальная проверка
                    setTimeout(function() {
                        console.log('Финальная проверка класса body:', document.body.className);
                        if (!document.body.className || document.body.className.indexOf('view-mode-') < 0) {
                            console.log('Класс не установлен, устанавливаем split...');
                            setViewMode('split');
                        }
                    }, 300);
                    
                    if (editor) {
                        editor.focus();
                        notifyChange();
                    }
                }
                
                // Запускаем инициализацию
                setTimeout(initializeApp, 300);

            </script>
        </body>
        </html>
    ";
        }

        private string BuildHtmlShell_Old() 
        {
            return @"<!DOCTYPE html>
                <html>
                <head>
                    <!-- Заголовок и мета-теги -->
                    <meta charset=\""utf-8\""/>
                    <meta http-equiv=\""X-UA-Compatible\"" content=\""IE=edge\"" />
                    <meta name=\""viewport\"" content=\""width=device-width, initial-scale=1\"" />
                    <title>Markdown Editor</title>
                        <!-- Стили -->
                        <style>
                             html, body { 
                                height: 100%; 
                                margin: 0; 
                                font-family: Segoe UI, Arial, sans-serif; 
                             }   
                             .container { 
                                display: flex; 
                                height: 100%; 
                             }
                             
                             #editor { 
                                width: 50%; 
                                height: 100%; 
                                border: none; 
                                padding: 12px; 
                                font-family: Consolas, monospace; 
                                font-size: 13px; 
                                box-sizing: border-box; 
                                outline: none; 
                                resize: none; 
                                border-right: 1px solid #ddd; 
                             }
                            
                             #preview { 
                                width: 50%; 
                                height: 100%; 
                                overflow: auto; 
                                padding: 16px; 
                                box-sizing: border-box; 
                             }

                             pre { 
                                background: #f6f8fa; 
                                padding: 10px; 
                                overflow: auto; 
                             }
                             code { 
                                font-family: Consolas, monospace; 
                             }
                            
                        </style>
                        
                </head>
                <body>
                    <!-- Структура редактора -->
                    <div class=\""container\"">
                        <textarea id=\""editor\"" placeholder=\""Введите Markdown...""></textarea>
                        <div id=\""preview\""></div>
                    </div>
                    <!-- Скрипты -->
                    <script src=\""https://cdn.jsdelivr.net/npm/dompurify@3.1.0/dist/purify.min.js\""></script>
                    <script src=\""https://cdn.jsdelivr.net/npm/prismjs@1.29.0/prism.min.js\""></script>
                    <script src=\""https://cdn.jsdelivr.net/npm/prismjs@1.29.0/plugins/autoloader/prism-autoloader.min.js\""></script>
                    <script>Prism.plugins.autoloader.languages_path = 'https://cdn.jsdelivr.net/npm/prismjs@1.29.0/components/';</script>
                    <script src=\""https://cdn.jsdelivr.net/npm/mermaid@10.9.0/dist/mermaid.min.js\""></script>
                    <script>mermaid.initialize({ startOnLoad: false, securityLevel: 'strict' });</script>
                    <script>window.MathJax = { tex: { inlineMath: [['$', '$'], ['\\\(', '\\\)']] }, svg: { fontCache: 'global' } };</script>
                    <script src=\""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js\""></script>
                    
                    <script>
                        // Базовые переменные
                        const editor = document.getElementById('editor');
                        const preview = document.getElementById('preview');
        
                        // Функция для отправки сообщений в C#
                        function postToHost(type, text) {
                            try {
                                const b64 = btoa(unescape(encodeURIComponent(text || '')));
                                if (window.chrome && window.chrome.webview) {
                                    window.chrome.webview.postMessage(type + '|' + b64);
                                }
                                 else if (window.chrome && window.chrome.webview) {
                                    window.chrome.webview.postMessage(type + '|' + b64);
                                }
                                else if (window.external && typeof window.external.notify === 'function') {
                                    window.external.notify(type + '|' + b64);
                                }
                                } catch(e) { 
                                    console.error('Ошибка отправки:', e); 
                                }
                        }
        
                        // Уведомление об изменениях с задержкой
                        function debounce(fn, ms) { 
                            let t; 
                            return function() { 
                                clearTimeout(t); 
                                t = setTimeout(() => fn.apply(this, arguments), ms); 
                            } 
                        }
        
                        function notifyChange() { 
                            postToHost('mdChanged', editor.value); 
                        }
        
                        // Слушаем изменения в редакторе
                        editor.addEventListener('input', debounce(notifyChange, 120));

                        // Методы для вызова из C#
                        window.editorSetValue = function(text) { 
                            editor.value = text || ''; 
                            notifyChange(); 
                        }
                        
                        window.editorGetValue = function() { 
                            return editor.value || ''; 
                        }

                        window.insertAroundSelection = function(prefix, suffix) {
                            prefix = prefix || ''; 
                            suffix = suffix || '';
                            const start = editor.selectionStart || 0;
                            const end = editor.selectionEnd || 0;
                            const val = editor.value;
            
                            editor.value = val.substring(0, start) + 
                                prefix + 
                                val.substring(start, end) + 
                                suffix + 
                                val.substring(end);
            
                            const newPos = start + prefix.length + (end - start);
                            editor.setSelectionRange(newPos, newPos);
                            editor.focus();
                            notifyChange();
                        }
                        
                         window.insertSnippet = function(snippet) {
                            const pos = editor.selectionStart || 0;
                            const val = editor.value;
                            editor.value = val.substring(0, pos) + snippet + val.substring(pos);
                            const newPos = pos + snippet.length;
                            editor.setSelectionRange(newPos, newPos);
                            editor.focus();
                            notifyChange();
                        }

                        window.renderHtml = function(html) {
                            try {
                                // Базовая очистка и отображение
                                const clean = DOMPurify.sanitize(html || '', { 
                                ADD_ATTR: ['target', 'rel', 'class', 'style', 'id'] 
                                });
                                preview.innerHTML = clean;
                                // Преобразовать блоки кода mermaid в divs
                                const mermaidBlocks = preview.querySelectorAll('pre code.language-mermaid');
                                mermaidBlocks.forEach(code => {
                                    const pre = code.parentElement;
                                    const wrapper = document.createElement('div');
                                    wrapper.className = 'mermaid';
                                    wrapper.textContent = code.textContent;
                                    pre.replaceWith(wrapper);
                                });
                                Prism.highlightAllUnder(preview);
                                if (window.mermaid) {
                                    mermaid.init(undefined, preview.querySelectorAll('.mermaid'));
                                }
                                if (window.MathJax && MathJax.typesetPromise) {
                                    MathJax.typesetPromise([preview]).catch(err => console.error(err));
                                }
                                window.addEventListener('load', function() {
                                    editor.focus();
                                    //notifyChange(); // Отправить начальное состояние
                                    setTimeout(() => notifyChange(), 100);
                                });

                            } catch(e) { 
                                console.error('Ошибка рендеринга:', e); 
                            }
                        }

                    </script>
                </body>
                </html>
            ";
        }

    }
}
