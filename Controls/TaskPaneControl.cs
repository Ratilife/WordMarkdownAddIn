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
            _renderer = new Services.MarkdownRenderService();   // Средство визуализации 
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
            System.Diagnostics.Debug.WriteLine("[C#] ✓ Обработчик WebMessageReceived подписан");
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            
            // Подписываемся на событие загрузки навигации для отладки
            _webView.CoreWebView2.NavigationCompleted += (navSender, navArgs) =>
            {
                System.Diagnostics.Debug.WriteLine("[C#] NavigationCompleted: HTML загружен");
                // Проверяем, что элементы существуют
                _webView.CoreWebView2.ExecuteScriptAsync(@"
                    setTimeout(function() {
                        var editor = document.getElementById('editor');
                        var preview = document.getElementById('preview');
                        var btnSplit = document.getElementById('btn-split');
                        var btnMarkdown = document.getElementById('btn-markdown');
                        var btnHtml = document.getElementById('btn-html');
                        var viewControls = document.querySelector('.view-controls');
                        
                        console.log('[JS] Проверка элементов после загрузки:');
                        console.log('  editor:', editor ? 'найден' : 'НЕ НАЙДЕН');
                        console.log('  preview:', preview ? 'найден' : 'НЕ НАЙДЕН');
                        console.log('  btnSplit:', btnSplit ? 'найден' : 'НЕ НАЙДЕН');
                        console.log('  btnMarkdown:', btnMarkdown ? 'найден' : 'НЕ НАЙДЕН');
                        console.log('  btnHtml:', btnHtml ? 'найден' : 'НЕ НАЙДЕН');
                        console.log('  viewControls:', viewControls ? 'найден' : 'НЕ НАЙДЕН');
                        console.log('  body.className:', document.body.className);
                        console.log('  body.style.display:', document.body.style.display);
                        console.log('  viewControls.style.display:', viewControls ? viewControls.style.display : 'N/A');
                    }, 500);
                ");
            };
            
            var html = BuildHtmlShell();
            System.Diagnostics.Debug.WriteLine($"[C#] Загрузка HTML, длина: {html.Length}");
            _webView.CoreWebView2.NavigateToString(html);
            System.Diagnostics.Debug.WriteLine("[C#] ✓ HTML загружен в WebView2");
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
                System.Diagnostics.Debug.WriteLine($"[C#] ===== WebMessageReceived вызван ===== ");
                System.Diagnostics.Debug.WriteLine($"[C#] Получено сообщение (первые 200 символов): {json?.Substring(0, Math.Min(200, json?.Length ?? 0))}...");
                System.Diagnostics.Debug.WriteLine($"[C#] Длина сообщения: {json?.Length ?? 0}");
                
                if (string.IsNullOrEmpty(json))
                {
                    System.Diagnostics.Debug.WriteLine("[C#] ✗ Сообщение пустое или null");
                    return;                                                 // Если сообщение пустое или null - выходим из метода
                }
                var parts = json.Split(new[] { '|' }, 2);                                               // Split - разделяет строку по символу | на 2 части
                                                                                                        // Пример: "mdChanged|SGVsbG8=" → ["mdChanged", "SGVsbG8="]
                if (parts.Length != 2)
                {
                    System.Diagnostics.Debug.WriteLine($"[C#] ✗ Неверный формат сообщения, parts.Length={parts.Length}");
                    System.Diagnostics.Debug.WriteLine($"[C#] Полное сообщение: {json}");
                    return;                                                          // Если сообщение не соответствует формату - игнорируем
                }
                var type = parts[0];                                                                    // тип сообщения (например: "mdChanged")
                System.Diagnostics.Debug.WriteLine($"[C#] ✓ Тип сообщения: {type}");
                
                try
                {
                    var payload = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1]));              // декодированные данные из Base64
                                                                                                        // Convert.FromBase64String() - преобразует Base64 в байты
                                                                                                        // Encoding.UTF8.GetString() - преобразует байты в строку UTF-8
                    System.Diagnostics.Debug.WriteLine($"[C#] ✓ Payload декодирован, длина: {payload.Length}");
                    
                    //Обработка изменения Markdown
                    if (type == "mdChanged")                                                                // Проверка типа - если это сообщение об изменении Markdown
                    {
                        System.Diagnostics.Debug.WriteLine($"[C#] ===== Обработка mdChanged ===== ");
                        _latestMarkdown = payload;                                                          // Сохранение - обновление кэша
                        System.Diagnostics.Debug.WriteLine($"[C#] Markdown сохранен в кэш, длина: {_latestMarkdown.Length}");
                        
                        var html = _renderer.RenderoHtml(payload);                                          // Конвертация - Markdown → HTML
                        System.Diagnostics.Debug.WriteLine($"[C#] ✓ Markdown конвертирован в HTML, длина: {html.Length}");
                        System.Diagnostics.Debug.WriteLine($"[C#] HTML (первые 200 символов): {html.Substring(0, Math.Min(200, html.Length))}...");
                        
                        PostRenderHtml(html);                                                               // Отправка - показ HTML в preview
                        System.Diagnostics.Debug.WriteLine($"[C#] ✓ PostRenderHtml вызван");
                        System.Diagnostics.Debug.WriteLine($"[C#] ===== mdChanged обработан успешно ===== ");
                    }
                    
                    // Обработка изменения режима отображения
                    if (type == "viewModeChanged")
                    {
                        System.Diagnostics.Debug.WriteLine($"[C#] Обработка viewModeChanged: {payload}");
                        // Сохраняем режим в настройках приложения
                        try
                        {
                            Settings.Default.ViewMode = payload;
                            Settings.Default.Save();
                            System.Diagnostics.Debug.WriteLine($"[C#] ✓ Режим сохранен: {payload}");
                        }
                        catch (Exception ex2)
                        {
                            System.Diagnostics.Debug.WriteLine($"[C#] ✗ Ошибка сохранения режима: {ex2.Message}");
                        }
                    }
                }
                catch (Exception decodeEx)
                {
                    System.Diagnostics.Debug.WriteLine($"[C#] ✗ Ошибка декодирования Base64: {decodeEx.Message}");
                    System.Diagnostics.Debug.WriteLine($"[C#] Base64 строка (первые 100 символов): {parts[1]?.Substring(0, Math.Min(100, parts[1]?.Length ?? 0))}...");
                }

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[C#] ===== ОШИБКА в WebMessageReceived ===== ");
                System.Diagnostics.Debug.WriteLine($"[C#] Ошибка обработки сообщения: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[C#] Stack trace: {ex.StackTrace}");
                System.Diagnostics.Debug.WriteLine($"[C#] ===== Конец ошибки ===== ");
            }
        }


        /// <summary>
        /// Асинхронно отображает HTML-содержимое в элементе управления WebView2.
        /// Метод проверяет готовность компонента WebView2, кодирует переданный HTML в Base64,
        /// передаёт его в JavaScript-функцию renderHtml, выполняемую внутри WebView2, для отображения.
        /// Обрабатывает возможные исключения, связанные с состоянием WebView2.
        /// </summary>
        /// <param name="html">Строка HTML-разметки, которая будет отображена.</param>
        /// <returns>Метод не возвращает значение (void). Вызов асинхронный.</returns>
        /// <remarks>
        /// Требует, чтобы _webView и его CoreWebView2 были инициализированы и готовы к взаимодействию.
        /// Использует JavaScript-функцию `window.renderHtml` и вспомогательную функцию `base64ToUtf8` в браузере.
        /// </remarks>
        private void PostRenderHtml(string html) 
        {
            if (!_coreReady || _webView == null)
            {
                System.Diagnostics.Debug.WriteLine("[PostRenderHtml] WebView2 не готов");
                return;                                                // Проверка готовности WebView2
            }
            try
            {
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null)
                {
                    System.Diagnostics.Debug.WriteLine("[PostRenderHtml] CoreWebView2 равен null");
                    return;
                }
                
                var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(html));                         // Кодирование HTML в Base64
                                                                                                        // Encoding.UTF8.GetBytes(html) - преобразует строку HTML в байты UTF-8
                                                                                                        // Convert.ToBase64String() - кодирует байты в строку Base64
                                                                                                        // Пример: <p>Hello</p> → "PHA+SGVsbG88L3A+"
                System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Отправка HTML длиной {html.Length} символов, Base64 длиной {b64.Length}");
                
                // Экранируем Base64 для безопасной передачи в JavaScript как строковый литерал
                var escapedB64 = b64.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", "\\r").Replace("\n", "\\n");
                
                var script = $@"
                    (function() {{
                        try {{
                            var b64 = '{escapedB64}';
                            
                            // Пробуем через функции
                            if (typeof window.renderHtml === 'function' && typeof window.base64ToUtf8 === 'function') {{
                                var html = window.base64ToUtf8(b64);
                                window.renderHtml(html);
                                return 'SUCCESS: via functions';
                            }}
                            
                            // Fallback: напрямую устанавливаем innerHTML
                            var preview = document.getElementById('preview');
                            if (preview) {{
                                // Простое декодирование Base64
                                var binary = atob(b64);
                                var bytes = new Uint8Array(binary.length);
                                for (var i = 0; i < binary.length; i++) {{
                                    bytes[i] = binary.charCodeAt(i);
                                }}
                                var html = new TextDecoder('utf-8').decode(bytes);
                                
                                // Устанавливаем HTML напрямую
                                if (typeof DOMPurify !== 'undefined') {{
                                    preview.innerHTML = DOMPurify.sanitize(html, {{ ADD_ATTR: ['target', 'rel', 'class', 'style', 'id'] }});
                                }} else {{
                                    preview.innerHTML = html;
                                }}
                                return 'SUCCESS: direct';
                            }}
                            
                            return 'ERROR: preview not found';
                        }} catch(e) {{
                            return 'ERROR: ' + e.message;
                        }}
                    }})();";
                
                // Выполняем скрипт
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Выполняем скрипт...");
                    var task = coreWebView2.ExecuteScriptAsync(script);
                    
                    // Получаем результат для отладки
                    task.ContinueWith(t =>
                    {
                        try
                        {
                            if (t.Status == TaskStatus.RanToCompletion && !t.IsFaulted && !t.IsCanceled)
                            {
                                var result = t.Result;
                                var resultStr = result?.ToString() ?? "null";
                                // Убираем кавычки из JSON строки
                                if (resultStr.StartsWith("\"") && resultStr.EndsWith("\""))
                                {
                                    resultStr = resultStr.Substring(1, resultStr.Length - 2);
                                }
                                System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Результат: {resultStr}");
                            }
                            else if (t.IsFaulted)
                            {
                                System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Ошибка: {t.Exception?.GetBaseException()?.Message}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Статус: {t.Status}");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Ошибка в ContinueWith: {ex.Message}");
                        }
                    }, TaskScheduler.Default);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Исключение: {ex.Message}");
                }
                                                                                                        // JSON.parse() безопасно парсит JSON строку, избегая проблем с XML/CDATA
                                                                                                        // window.renderHtml() - вызов JavaScript функции из C#
                                                                                                        // base64ToUtf8() - правильное декодирование Base64 → UTF-8
                                                                                                        // void(0); - предотвращает возврат значения (оптимизация)
            }
            catch (InvalidCastException ex)
            {
                // WebView2 еще не готов, игнорируем
                System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] InvalidCastException: {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                // Другие ошибки тоже игнорируем
                System.Diagnostics.Debug.WriteLine($"[PostRenderHtml] Exception: {ex.Message}");
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
                // Экранируем Base64 для безопасной передачи
                var escapedB64 = b64.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", "\\r").Replace("\n", "\\n");
                var script = $@"
                    (function() {{
                        try {{
                            var b64 = '{escapedB64}';
                            window.editorSetValue(base64ToUtf8(b64));
                        }} catch(e) {{
                            console.error('[C#->JS] Ошибка в editorSetValue:', e);
                        }}
                    }})();";
                coreWebView2.ExecuteScriptAsync(script);                                                               // Отправка в JavaScript
                                                                                                                        // ExecuteScriptAsync() - выполняет JavaScript код асинхронно
                                                                                                                        // window.editorSetValue() - JS функция для установки значения редактора
                                                                                                                        // base64ToUtf8() - правильное декодирование Base64 → UTF-8
                                                                                                                        // notifyChange() вызывается внутри editorSetValue для обновления HTML
                
                // Явно конвертируем markdown в HTML и отправляем в preview
                // Это гарантирует, что HTML обновится даже если notifyChange не сработает
                var html = _renderer.RenderoHtml(_latestMarkdown);
                System.Diagnostics.Debug.WriteLine($"[SetMarkdown] Конвертация markdown длиной {_latestMarkdown.Length} в HTML длиной {html.Length}");
                PostRenderHtml(html);
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

        public WebView2 GetWebView()
        {
            return _webView;
        }

        //Метод предоставляет надежный способ получения актуального Markdown-текста,используя кэш, но при необходимости запрашивая данные из JavaScript.
        /// <summary>
        /// Асинхронно возвращает Markdown-содержимое редактора.
        /// Сначала пытается получить значение из кэша (<see cref="_latestMarkdown"/>).
        /// Если кэш пуст или не инициализирован, и компонент WebView2 готов,
        /// делает асинхронный запрос к JavaScript-редактору через <c>window.editorGetValue()</c>,
        /// обрабатывает результат и возвращает его.
        /// </summary>
        /// <returns>Задача (Task), результатом которой является строка Markdown-содержимого.
        /// Возвращает пустую строку, если кэш пуст и запрос к JavaScript не удался или WebView2 не готов.</returns>
        /// <remarks>
        /// Использует кэшированное значение, обновляемое, например, через <c>mdChanged</c>, для оптимизации.
        /// В случае ошибки возвращается пустая строка.
        /// </remarks>
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
        /// <summary>
        /// Удаляет внешние кавычки из строки JSON и декодирует экранированные управляющие символы.
        /// </summary>
        /// <param name="jsonQuoted">Строка, представляющая собой JSON-значение (обычно строку), полученное, например, из JavaScript.</param>
        /// <returns>Обработанная строка без внешних кавычек и с восстановленными управляющими символами (\n, \t, \", и т.д.).
        /// Возвращает пустую строку, если входное значение null или пустое.</returns>
        /// <remarks>
        /// Используется для обработки строк, полученных из <c>WebView2.ExecuteScriptAsync</c>, 
        /// так как результат выполнения JavaScript-кода возвращается в формате JSON.
        /// </remarks>
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
        /// <summary>
        /// Асинхронно вставляет указанные строки префикса и суффикса вокруг текущего выделения в редакторе, размещенном в WebView2.
        /// Использует JavaScript-функцию <c>window.insertAroundSelection</c> для выполнения вставки на стороне редактора.
        /// Префиксы и суффиксы передаются в JavaScript в виде закодированных в Base64 строк UTF-8, чтобы избежать проблем с экранированием.
        /// </summary>
        /// <param name="prefix">Строка, которая будет вставлена перед выделенным текстом. Может быть null или пустой.</param>
        /// <param name="suffix">Строка, которая будет вставлена после выделенного текста. Может быть null или пустой.</param>
        /// <remarks>
        /// Метод ничего не делает, если компонент WebView2 не инициализирован или недоступен.
        /// Все возникающие исключения игнорируются.
        /// </remarks>
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
                coreWebView2.ExecuteScriptAsync($"window.insertAroundSelection(base64ToUtf8('{p}'), base64ToUtf8('{s}'));void(0);"); // Отправка в JavaScript
                                                                                                                                // ExecuteScriptAsync() - выполняет JS код асинхронно
                                                                                                                                // window.insertAroundSelection() - JS функция для вставки вокруг выделения
                                                                                                                                // base64ToUtf8('{p}') - правильное декодирование Base64 → UTF-8 префикса
                                                                                                                                // base64ToUtf8('{s}') - правильное декодирование Base64 → UTF-8 суффикса
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
        /// <summary>
        /// Асинхронно вставляет указанный текстовый фрагмент (сниппет) в редактор, размещенный в WebView2.
        /// Использует JavaScript-функцию <c>window.insertSnippet</c> для выполнения вставки.
        /// Сниппет передаётся в JavaScript в виде закодированной в Base64 строки UTF-8.
        /// </summary>
        /// <param name="snippet">Текст сниппета для вставки. Может быть null или пустой строкой.</param>
        /// <remarks>
        /// Метод ничего не делает, если компонент WebView2 не инициализирован или недоступен.
        /// Все возникающие исключения игнорируются.
        /// </remarks>
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
                coreWebView2.ExecuteScriptAsync($"window.insertSnippet(base64ToUtf8('{b64}'));void(0);");                        // Отправка в JavaScript
                                                                                                                                // ExecuteScriptAsync() - выполняет JS код асинхронно
                                                                                                                                // window.insertSnippet() - JS функция для вставки сниппета
                                                                                                                                // base64ToUtf8('{b64}') - правильное декодирование Base64 → UTF-8
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
                    padding: 0;
                    font-family: Segoe UI, Arial, sans-serif; 
                    overflow: hidden;
                    display: flex;
                    flex-direction: column;
                }   
                
                /* Панель управления режимами */
                .view-controls {
                    display: flex !important;
                    flex-shrink: 0;
                    gap: 4px;
                    padding: 8px;
                    background: #f5f5f5;
                    border-bottom: 1px solid #ddd;
                    min-height: 40px;
                    box-sizing: border-box;
                    z-index: 10;
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
                    flex: 1;
                    min-height: 0;
                    overflow: hidden;
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
            <link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/prismjs@1.29.0/themes/prism-okaidia.css"" />
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
                // Базовые переменные (будут инициализированы после загрузки DOM)
                let editor = null;
                let preview = null;
                
                // Переменная для текущего режима отображения
                let currentViewMode = 'split'; // по умолчанию
                let rawHtml = ''; // Храним исходный HTML для режима HTML
                
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
                        
                        // Обновляем отображение при переключении режима
                        updatePreviewDisplay();
                        
                        // При переключении на HTML режим, если есть текст в редакторе, обновляем HTML
                        if (mode === 'html' && editor && editor.value && editor.value.trim()) {
                            // Отправляем текущий Markdown для конвертации в HTML
                            setTimeout(function() {
                                notifyChange();
                            }, 100);
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
                        // Правильное кодирование UTF-8 в Base64 для всех символов
                        function utf8ToBase64(str) {
                            return btoa(unescape(encodeURIComponent(str || '')));
                        }
                        const b64 = utf8ToBase64(text || '');
                        const message = type + '|' + b64;
                        console.log('[JS] ===== postToHost вызван =====');
                        console.log('[JS] postToHost: type=' + type + ', text length=' + (text || '').length + ', b64 length=' + b64.length);
                        
                        // Пробуем разные способы отправки сообщений
                        let sent = false;
                        
                        // Способ 1: window.chrome.webview.postMessage (WebView2) - основной способ
                        if (window.chrome && window.chrome.webview && typeof window.chrome.webview.postMessage === 'function') {
                            try {
                                window.chrome.webview.postMessage(message);
                                console.log('[JS] ✓ Сообщение отправлено через chrome.webview.postMessage');
                                sent = true;
                            } catch(e) {
                                console.error('[JS] ✗ Ошибка при отправке через chrome.webview.postMessage:', e);
                                console.error('[JS] Ошибка детали:', e.message, e.stack);
                            }
                        } else {
                            console.warn('[JS] window.chrome.webview.postMessage недоступен');
                            console.log('[JS] window.chrome:', window.chrome);
                            console.log('[JS] window.chrome.webview:', window.chrome ? window.chrome.webview : 'undefined');
                        }
                        
                        // Способ 2: window.external.notify (старый способ, для совместимости)
                        if (!sent && window.external && typeof window.external.notify === 'function') {
                            try {
                                window.external.notify(message);
                                console.log('[JS] ✓ Сообщение отправлено через window.external.notify');
                                sent = true;
                            } catch(e) {
                                console.error('[JS] ✗ Ошибка при отправке через window.external.notify:', e);
                            }
                        }
                        
                        if (!sent) {
                            console.error('[JS] ✗ Не найден способ отправки сообщения!');
                            console.error('[JS] Проверка доступности API:');
                            console.error('[JS]   window.chrome:', typeof window.chrome);
                            console.error('[JS]   window.chrome.webview:', window.chrome ? typeof window.chrome.webview : 'undefined');
                            console.error('[JS]   window.external:', typeof window.external);
                        } else {
                            console.log('[JS] ===== postToHost завершен успешно =====');
                        }
                    } catch(e) { 
                        console.error('[JS] ===== ОШИБКА в postToHost =====');
                        console.error('[JS] Ошибка отправки:', e); 
                        console.error('[JS] Сообщение:', e.message);
                        console.error('[JS] Stack trace:', e.stack);
                        console.error('[JS] ===== Конец ошибки =====');
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
                    if (!editor) {
                        console.error('[JS] notifyChange: editor не инициализирован');
                        // Пытаемся получить editor еще раз
                        editor = document.getElementById('editor');
                        if (!editor) {
                            console.error('[JS] notifyChange: элемент editor не найден в DOM');
                            return;
                        }
                    }
                    const markdownValue = editor.value || '';
                    console.log('[JS] ===== notifyChange вызван =====');
                    console.log('[JS] notifyChange: длина markdown:', markdownValue.length);
                    console.log('[JS] notifyChange: содержимое markdown (первые 100 символов):', markdownValue.substring(0, 100));
                    
                    // Отправляем сообщение в C#
                    postToHost('mdChanged', markdownValue);
                    
                    console.log('[JS] ===== notifyChange завершен =====');
                }

                // Слушаем изменения в редакторе (будет установлено после инициализации)
                function setupEditorListeners() {
                    if (!editor) {
                        console.error('[JS] setupEditorListeners: editor не найден, пытаемся получить...');
                        editor = document.getElementById('editor');
                    }
                    if (editor) {
                        // Удаляем старые обработчики, если они есть
                        // Создаем именованную функцию для обработчика, чтобы можно было удалить
                        if (editor._notifyChangeHandler) {
                            editor.removeEventListener('input', editor._notifyChangeHandler);
                        }
                        if (editor._keyupHandler) {
                            editor.removeEventListener('keyup', editor._keyupHandler);
                        }
                        
                        // Создаем обработчик с debounce
                        const debouncedNotify = debounce(notifyChange, 120);
                        editor._notifyChangeHandler = debouncedNotify;
                        
                        // Добавляем обработчик input для отслеживания изменений
                        editor.addEventListener('input', editor._notifyChangeHandler);
                        
                        // Также добавляем обработчик keyup для немедленного обновления при некоторых действиях
                        editor._keyupHandler = function() {
                            // Вызываем notifyChange без debounce для некоторых клавиш
                            notifyChange();
                        };
                        editor.addEventListener('keyup', editor._keyupHandler);
                        
                        console.log('[JS] Обработчики input и keyup установлены для editor');
                        
                        // Тестовый вызов для проверки
                        console.log('[JS] Тестовый вызов notifyChange...');
                        setTimeout(function() {
                            if (editor && editor.value) {
                                console.log('[JS] Editor имеет значение, вызываем notifyChange');
                                notifyChange();
                            }
                        }, 500);
                    } else {
                        console.error('[JS] Не удалось установить обработчик input: editor не найден');
                    }
                }

                // Методы для вызова из C#
                window.editorSetValue = function(text) { 
                    if (!editor) {
                        console.error('[JS] editorSetValue: editor не инициализирован');
                        // Пытаемся получить элемент еще раз
                        editor = document.getElementById('editor');
                        if (!editor) {
                            console.error('[JS] editorSetValue: элемент editor не найден в DOM');
                            return;
                        }
                    }
                    editor.value = text || ''; 
                    notifyChange(); 
                }
                
                window.editorGetValue = function() { 
                    if (!editor) {
                        console.error('[JS] editorGetValue: editor не инициализирован');
                        // Пытаемся получить элемент еще раз
                        editor = document.getElementById('editor');
                        if (!editor) {
                            console.error('[JS] editorGetValue: элемент editor не найден в DOM');
                            return '';
                        }
                    }
                    return editor.value || ''; 
                }

                window.insertAroundSelection = function(prefix, suffix) {
                    if (!editor) {
                        console.error('[JS] insertAroundSelection: editor не инициализирован');
                        editor = document.getElementById('editor');
                        if (!editor) return;
                    }
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
                    if (!editor) {
                        console.error('[JS] insertSnippet: editor не инициализирован');
                        editor = document.getElementById('editor');
                        if (!editor) return;
                    }
                    const pos = editor.selectionStart || 0;
                    const val = editor.value;
                    editor.value = val.substring(0, pos) + snippet + val.substring(pos);
                    const newPos = pos + snippet.length;
                    editor.setSelectionRange(newPos, newPos);
                    editor.focus();
                    notifyChange();
                }

                // Функция обновления отображения в зависимости от режима
                function updatePreviewDisplay() {
                    try {
                        console.log('[JS] updatePreviewDisplay: начало');
                        console.log('[JS] updatePreviewDisplay: rawHtml length:', (rawHtml || '').length);
                        
                        // Проверяем, что preview существует
                        if (!preview) {
                            console.error('[JS] updatePreviewDisplay: preview элемент не найден, пытаемся получить...');
                            preview = document.getElementById('preview');
                            if (!preview) {
                                console.error('[JS] updatePreviewDisplay: preview все еще не найден!');
                                return;
                            }
                        }
                        
                        // Во всех режимах рендерим HTML одинаково (отформатированный контент)
                        const htmlContent = rawHtml || '';
                        console.log('[JS] updatePreviewDisplay: htmlContent length:', htmlContent.length);
                        console.log('[JS] updatePreviewDisplay: htmlContent (первые 100 символов):', htmlContent.substring(0, 100));
                        
                        // Если HTML пустой, просто очищаем preview
                        if (!htmlContent.trim()) {
                            console.log('[JS] updatePreviewDisplay: HTML пустой, очищаем preview');
                            preview.innerHTML = '';
                            return;
                        }
                        
                        // Проверяем наличие DOMPurify
                        if (typeof DOMPurify === 'undefined') {
                            console.error('[JS] updatePreviewDisplay: DOMPurify не загружен! Используем прямое присваивание.');
                            preview.innerHTML = htmlContent;
                        } else {
                            const clean = DOMPurify.sanitize(htmlContent, { 
                                ADD_ATTR: ['target', 'rel', 'class', 'style', 'id'] 
                            });
                            preview.innerHTML = clean;
                            console.log('[JS] updatePreviewDisplay: preview.innerHTML обновлен через DOMPurify, длина:', clean.length);
                        }
                        
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
                        console.error('[JS] Ошибка в updatePreviewDisplay:', e);
                    }
                }
                
                // Функция для экранирования HTML в тексте
                function escapeHtml(text) {
                    const div = document.createElement('div');
                    div.textContent = text;
                    return div.innerHTML;
                }
                
                // Функция для правильного декодирования Base64 → UTF-8
                function base64ToUtf8(base64) {
                    try {
                        // Декодируем Base64 в бинарную строку (Latin-1)
                        const binary = atob(base64);
                        // Создаём массив байтов
                        const bytes = new Uint8Array(binary.length);
                        for (let i = 0; i < binary.length; i++) {
                            bytes[i] = binary.charCodeAt(i);
                        }
                        // Декодируем байты как UTF-8
                        return new TextDecoder('utf-8').decode(bytes);
                    } catch(e) {
                        console.error('Ошибка декодирования Base64:', e);
                        // Fallback: пробуем простой atob (для совместимости)
                        try {
                            return decodeURIComponent(escape(atob(base64)));
                        } catch(e2) {
                            console.error('Ошибка fallback декодирования:', e2);
                            return '';
                        }
                    }
                }
                
                // Делаем функцию доступной глобально
                window.base64ToUtf8 = base64ToUtf8;
                
                window.renderHtml = function(html) {
                    try {
                        console.log('[JS] ===== renderHtml вызван =====');
                        console.log('[JS] renderHtml: длина HTML:', (html || '').length);
                        console.log('[JS] renderHtml: preview существует:', !!preview);
                        console.log('[JS] renderHtml: DOMPurify существует:', typeof DOMPurify !== 'undefined');
                        
                        if (!html) {
                            console.warn('[JS] renderHtml получил пустой HTML');
                            html = '';
                        }
                        
                        // Проверяем, что preview существует
                        if (!preview) {
                            console.error('[JS] renderHtml: preview не найден, пытаемся получить...');
                            preview = document.getElementById('preview');
                            if (!preview) {
                                console.error('[JS] renderHtml: preview все еще не найден после попытки получения!');
                                return;
                            }
                        }
                        
                        // Сохраняем исходный HTML
                        rawHtml = html;
                        console.log('[JS] rawHtml сохранен, длина:', rawHtml.length);
                        console.log('[JS] rawHtml содержимое (первые 100 символов):', rawHtml.substring(0, 100));
                        
                        // Обновляем отображение в зависимости от текущего режима
                        console.log('[JS] Вызываем updatePreviewDisplay...');
                        updatePreviewDisplay();
                        console.log('[JS] updatePreviewDisplay завершен');
                        console.log('[JS] preview.innerHTML длина после обновления:', preview.innerHTML.length);
                        console.log('[JS] ===== renderHtml завершен успешно =====');
                    } catch(e) { 
                        console.error('[JS] ===== ОШИБКА в renderHtml =====');
                        console.error('[JS] Ошибка рендеринга:', e); 
                        console.error('[JS] Сообщение:', e.message);
                        console.error('[JS] Stack trace:', e.stack);
                        console.error('[JS] ===== Конец ошибки =====');
                    }
                }

                // Инициализация после загрузки
                function initializeApp() {
                    // Проверяем готовность DOM
                    if (!document.body) {
                        console.log('[JS] DOM еще не готов, повтор через 200мс...');
                        setTimeout(initializeApp, 200);
                        return;
                    }
                    
                    console.log('[JS] DOM готов, начинаем инициализацию');
                    
                    // Инициализируем переменные editor и preview после загрузки DOM
                    editor = document.getElementById('editor');
                    preview = document.getElementById('preview');
                    
                    console.log('[JS] Элементы найдены:');
                    console.log('  editor:', editor ? 'найден' : 'НЕ НАЙДЕН');
                    console.log('  preview:', preview ? 'найден' : 'НЕ НАЙДЕН');
                    
                    // Проверяем кнопки
                    var btnSplit = document.getElementById('btn-split');
                    var btnMarkdown = document.getElementById('btn-markdown');
                    var btnHtml = document.getElementById('btn-html');
                    var viewControls = document.querySelector('.view-controls');
                    
                    console.log('  btnSplit:', btnSplit ? 'найден' : 'НЕ НАЙДЕН');
                    console.log('  btnMarkdown:', btnMarkdown ? 'найден' : 'НЕ НАЙДЕН');
                    console.log('  btnHtml:', btnHtml ? 'найден' : 'НЕ НАЙДЕН');
                    console.log('  viewControls:', viewControls ? 'найден' : 'НЕ НАЙДЕН');
                    
                    if (!editor) {
                        console.error('[JS] Элемент editor не найден!');
                    }
                    if (!preview) {
                        console.error('[JS] Элемент preview не найден!');
                    }
                    if (!viewControls) {
                        console.error('[JS] Элемент view-controls не найден!');
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
                    
                    // Устанавливаем режим (гарантируем, что preview будет виден)
                    console.log('[JS] Устанавливаем режим:', savedMode);
                    setViewMode(savedMode);
                    console.log('[JS] Режим установлен, класс body:', document.body.className);

                    // Инициализируем кнопки еще раз (на всякий случай)
                    initViewModeButtons();
                    
                    // Финальная проверка - убеждаемся, что режим split установлен и preview виден
                    setTimeout(function() {
                        console.log('[JS] Финальная проверка класса body:', document.body.className);
                        console.log('[JS] Финальная проверка preview:', preview ? 'найден' : 'НЕ НАЙДЕН');
                        if (preview) {
                            console.log('[JS] preview.style.display:', preview.style.display);
                            console.log('[JS] preview.offsetWidth:', preview.offsetWidth);
                            console.log('[JS] preview.offsetHeight:', preview.offsetHeight);
                        }
                        if (!document.body.className || document.body.className.indexOf('view-mode-') < 0) {
                            console.log('[JS] Класс не установлен, устанавливаем split...');
                            setViewMode('split');
                        }
                        // Принудительно показываем preview если он скрыт
                        if (preview && preview.offsetWidth === 0) {
                            console.log('[JS] Preview скрыт, принудительно показываем');
                            preview.style.display = 'block';
                            preview.style.width = '50%';
                        }
                        
                        // Тестовая отправка сообщения для проверки связи JavaScript → C#
                        console.log('[JS] ===== Тестовая отправка сообщения в C# =====');
                        postToHost('mdChanged', editor ? (editor.value || '') : '');
                        console.log('[JS] ===== Тестовая отправка завершена =====');
                    }, 500);
                    
                    // Устанавливаем обработчики событий для editor
                    setupEditorListeners();
                    
                    if (editor) {
                        editor.focus();
                        // Отправляем начальное состояние markdown для конвертации в HTML
                        // Это гарантирует, что preview будет обновлен при загрузке
                        console.log('[JS] Отправка начального состояния markdown для конвертации в HTML');
                        console.log('[JS] Начальное значение editor:', editor.value ? editor.value.substring(0, 50) : '(пусто)');
                        
                        // Проверяем, что обработчик установлен
                        setTimeout(function() {
                            if (editor) {
                                // Проверяем наличие обработчика
                                var hasInputHandler = editor._notifyChangeHandler !== undefined;
                                console.log('[JS] Проверка обработчика input:', hasInputHandler ? 'установлен' : 'НЕ установлен');
                                
                                // Если обработчик не установлен, устанавливаем его снова
                                if (!hasInputHandler) {
                                    console.warn('[JS] Обработчик input не найден, устанавливаем заново...');
                                    setupEditorListeners();
                                }
                                
                                // Отправляем начальное состояние
                                notifyChange();
                            }
                        }, 300);
                    } else {
                        console.error('[JS] Не удалось получить элемент editor для focus');
                    }
                }
                
                // Запускаем инициализацию
                // Используем DOMContentLoaded для гарантии готовности DOM
                if (document.readyState === 'loading') {
                    document.addEventListener('DOMContentLoaded', function() {
                        console.log('[JS] DOMContentLoaded сработал, запускаем initializeApp');
                        setTimeout(initializeApp, 100);
                    });
                } else {
                    console.log('[JS] DOM уже готов, запускаем initializeApp');
                    setTimeout(initializeApp, 100);
                }

            </script>
        </body>
        </html>
    ";
        }

        /// <summary>
        /// Восстанавливает HTML оболочку панели после экспорта Mermaid диаграмм.
        /// Восстанавливает markdown содержимое и переключает в HTML режим, сохраняя кнопки переключения режимов.
        /// </summary>
        public async Task RestoreHtmlShellAsync()
        {
            if (!_coreReady || _webView == null) return;
            
            try
            {
                var coreWebView2 = _webView.CoreWebView2;
                if (coreWebView2 == null) return;
                
                // Сохраняем текущий markdown перед восстановлением
                string savedMarkdown = _latestMarkdown;
                
                // Создаем TaskCompletionSource для ожидания завершения навигации
                var navigationTcs = new TaskCompletionSource<bool>();
                bool navigationCompleted = false;
                
                // Подписываемся на событие завершения навигации
                void NavigationHandler(object sender, CoreWebView2NavigationCompletedEventArgs e)
                {
                    if (!navigationCompleted)
                    {
                        navigationCompleted = true;
                        coreWebView2.NavigationCompleted -= NavigationHandler;
                        navigationTcs.SetResult(true);
                    }
                }
                
                coreWebView2.NavigationCompleted += NavigationHandler;
                
                // Восстанавливаем HTML оболочку
                var html = BuildHtmlShell();
                coreWebView2.NavigateToString(html);
                
                // Ждем завершения навигации (с таймаутом 5 секунд)
                await Task.WhenAny(navigationTcs.Task, Task.Delay(5000));
                
                if (!navigationCompleted)
                {
                    coreWebView2.NavigationCompleted -= NavigationHandler;
                    System.Diagnostics.Debug.WriteLine("[RestoreHtmlShellAsync] Таймаут ожидания навигации");
                }
                
                // Дополнительная задержка для инициализации JavaScript
                await Task.Delay(1000);
                
                // Проверяем, что элементы загружены
                var checkScript = @"
                    (function() {
                        var viewControls = document.querySelector('.view-controls');
                        var editor = document.getElementById('editor');
                        return (viewControls && editor) ? 'ready' : 'not ready';
                    })();
                ";
                var checkResult = await coreWebView2.ExecuteScriptAsync(checkScript);
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] Проверка готовности: {checkResult}");
                
                // Восстанавливаем markdown содержимое
                if (!string.IsNullOrEmpty(savedMarkdown))
                {
                    SetMarkdown(savedMarkdown);
                    await Task.Delay(800);
                }
                
                // Переключаем в HTML режим
                SetViewMode("html");
                
                // Функция для гарантированного отображения кнопок
                string ensureButtonsVisibleScript = @"
                    (function() {
                        var viewControls = document.querySelector('.view-controls');
                        if (viewControls) {
                            viewControls.style.display = 'flex';
                            viewControls.style.visibility = 'visible';
                            viewControls.style.opacity = '1';
                            viewControls.style.height = 'auto';
                            viewControls.style.minHeight = '40px';
                        }
                        var btnSplit = document.getElementById('btn-split');
                        var btnMarkdown = document.getElementById('btn-markdown');
                        var btnHtml = document.getElementById('btn-html');
                        if (btnSplit) {
                            btnSplit.style.display = 'block';
                            btnSplit.style.visibility = 'visible';
                            btnSplit.style.opacity = '1';
                        }
                        if (btnMarkdown) {
                            btnMarkdown.style.display = 'block';
                            btnMarkdown.style.visibility = 'visible';
                            btnMarkdown.style.opacity = '1';
                        }
                        if (btnHtml) {
                            btnHtml.style.display = 'block';
                            btnHtml.style.visibility = 'visible';
                            btnHtml.style.opacity = '1';
                        }
                        // Переинициализируем обработчики кнопок, если функция существует
                        if (typeof initViewModeButtons === 'function') {
                            try {
                                initViewModeButtons();
                            } catch(e) {
                                console.error('[RestoreHtmlShell] Ошибка инициализации кнопок:', e);
                            }
                        }
                        var allFound = viewControls && btnSplit && btnMarkdown && btnHtml;
                        console.log('[RestoreHtmlShell] Кнопки восстановлены:', {
                            viewControls: viewControls ? 'найден' : 'НЕ НАЙДЕН',
                            btnSplit: btnSplit ? 'найден' : 'НЕ НАЙДЕН',
                            btnMarkdown: btnMarkdown ? 'найден' : 'НЕ НАЙДЕН',
                            btnHtml: btnHtml ? 'найден' : 'НЕ НАЙДЕН',
                            allFound: allFound
                        });
                        return allFound ? 'success' : 'failed';
                    })();
                ";
                
                // Первая попытка - сразу после установки режима
                await Task.Delay(500);
                var result1 = await coreWebView2.ExecuteScriptAsync(ensureButtonsVisibleScript);
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] Первая попытка восстановления кнопок: {result1}");
                
                // Вторая попытка - через дополнительную задержку
                await Task.Delay(500);
                var result2 = await coreWebView2.ExecuteScriptAsync(ensureButtonsVisibleScript);
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] Вторая попытка восстановления кнопок: {result2}");
                
                // Третья попытка - финальная проверка
                await Task.Delay(300);
                var result3 = await coreWebView2.ExecuteScriptAsync(ensureButtonsVisibleScript);
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] Третья попытка восстановления кнопок: {result3}");
                
                // Финальная проверка видимости кнопок
                var finalCheck = await coreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var viewControls = document.querySelector('.view-controls');
                        var btnSplit = document.getElementById('btn-split');
                        var btnMarkdown = document.getElementById('btn-markdown');
                        var btnHtml = document.getElementById('btn-html');
                        var allVisible = viewControls && 
                                        viewControls.style.display !== 'none' &&
                                        btnSplit && btnSplit.style.display !== 'none' &&
                                        btnMarkdown && btnMarkdown.style.display !== 'none' &&
                                        btnHtml && btnHtml.style.display !== 'none';
                        console.log('[RestoreHtmlShell] Финальная проверка:', {
                            viewControlsVisible: viewControls && viewControls.style.display !== 'none',
                            btnSplitVisible: btnSplit && btnSplit.style.display !== 'none',
                            btnMarkdownVisible: btnMarkdown && btnMarkdown.style.display !== 'none',
                            btnHtmlVisible: btnHtml && btnHtml.style.display !== 'none',
                            allVisible: allVisible
                        });
                        return allVisible ? 'all_visible' : 'some_hidden';
                    })();
                ");
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] Финальная проверка видимости: {finalCheck}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] Ошибка: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[RestoreHtmlShellAsync] StackTrace: {ex.StackTrace}");
            }
        }

       
    }
}
