# Отчет по работе модуля TaskPaneControl.cs

## Содержание
1. [Общая архитектура](#общая-архитектура)
2. [Инициализация системы](#инициализация-системы)
3. [Взаимодействие C# → JavaScript](#взаимодействие-c--javascript)
4. [Взаимодействие JavaScript → C#](#взаимодействие-javascript--c)
5. [Формат передачи данных](#формат-передачи-данных)
6. [Потоки данных](#потоки-данных)
7. [Основные функции и методы](#основные-функции-и-методы)
8. [Примеры взаимодействия](#примеры-взаимодействия)

---

## Общая архитектура

### Компоненты системы

```
┌─────────────────────────────────────────────────────────────┐
│                    Word Add-In (C#)                        │
│  ┌──────────────────────────────────────────────────────┐  │
│  │         TaskPaneControl (UserControl)                │  │
│  │  ┌──────────────────────────────────────────────┐   │  │
│  │  │          WebView2 Control                    │   │  │
│  │  │  ┌────────────────────────────────────────┐  │   │  │
│  │  │  │      HTML + JavaScript                  │  │   │  │
│  │  │  │  ┌──────────────────────────────────┐ │  │   │  │
│  │  │  │  │  Markdown Editor (textarea)       │ │  │   │  │
│  │  │  │  └──────────────────────────────────┘ │  │   │  │
│  │  │  │  ┌──────────────────────────────────┐ │  │   │  │
│  │  │  │  │  HTML Preview (div#preview)      │ │  │   │  │
│  │  │  │  └──────────────────────────────────┘ │  │   │  │
│  │  │  └────────────────────────────────────────┘  │   │  │
│  │  └──────────────────────────────────────────────┘   │  │
│  │  ┌──────────────────────────────────────────────┐   │  │
│  │  │  MarkdownRenderService (C#)                  │   │  │
│  │  │  (Конвертация Markdown → HTML)               │   │  │
│  │  └──────────────────────────────────────────────┘   │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### Технологический стек

- **C#**: Backend логика, обработка данных, интеграция с Word
- **WebView2**: Контейнер для веб-контента (Microsoft Edge Chromium)
- **HTML/CSS**: Интерфейс пользователя (редактор + preview)
- **JavaScript**: Логика редактора, обработка событий, коммуникация с C#

---

## Инициализация системы

### Последовательность инициализации

```12:48:Controls/TaskPaneControl.cs
public class TaskPaneControl: UserControl
{
    private readonly WebView2 _webView;
    private readonly Services.MarkdownRenderService _renderer;
    private string _latestMarkdown = string.Empty;
    private bool _coreReady = false;

    public TaskPaneControl() 
    {
        _renderer = new Services.MarkdownRenderService();
        _webView = new WebView2
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(_webView);
        Load += OnLoadAsync;
    }

    private async void OnLoadAsync(object sender, EventArgs e) 
    {
        await _webView.EnsureCoreWebView2Async();
        _coreReady = true;
        _webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
        _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
        _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
        _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
        _webView.CoreWebView2.NavigateToString(BuildHtmlShell());
    }
```

**Шаги инициализации:**

1. **Создание компонентов** (конструктор):
   - Создается `MarkdownRenderService` для конвертации Markdown → HTML
   - Создается `WebView2` и добавляется на форму
   - Подписывается на событие `Load`

2. **Инициализация WebView2** (`OnLoadAsync`):
   - `EnsureCoreWebView2Async()` - асинхронная инициализация движка WebView2
   - Устанавливается флаг `_coreReady = true`
   - Подписывается обработчик `CoreWebView2_WebMessageReceived` для получения сообщений из JS
   - Загружается HTML через `NavigateToString(BuildHtmlShell())`

3. **Загрузка HTML** (`BuildHtmlShell`):
   - Генерируется полный HTML документ с встроенными CSS и JavaScript
   - HTML содержит:
     - Структуру редактора (textarea + preview div)
     - Стили для режимов отображения (split/markdown/html)
     - JavaScript функции для работы с редактором
     - Библиотеки: DOMPurify, Prism.js, Mermaid, MathJax

---

## Взаимодействие C# → JavaScript

### Механизм вызова

C# вызывает JavaScript функции через метод `ExecuteScriptAsync()`:

```csharp
coreWebView2.ExecuteScriptAsync("window.functionName(parameters);void(0);");
```

### Основные методы C# → JS

#### 1. Установка Markdown текста

```153:186:Controls/TaskPaneControl.cs
public void SetMarkdown(string markdown) 
{
    _latestMarkdown = markdown ?? string.Empty;
    if (!_coreReady || _webView == null) return;
    try
    {
        var coreWebView2 = _webView.CoreWebView2;
        if (coreWebView2 == null) return;
        
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(_latestMarkdown));
        coreWebView2.ExecuteScriptAsync($"window.editorSetValue(atob('{b64}'));void(0);");
    }
    catch (InvalidCastException)
    {
        return;
    }
    catch (Exception)
    {
        return;
    }
}
```

**Процесс:**
1. C# кодирует Markdown в Base64
2. Вызывает `window.editorSetValue(atob('Base64String'))`
3. JavaScript декодирует Base64 и устанавливает значение в textarea

#### 2. Отправка HTML для preview

```108:137:Controls/TaskPaneControl.cs
private void PostRenderHtml(string html) 
{
    if (!_coreReady || _webView == null) return;
    try
    {
        var coreWebView2 = _webView.CoreWebView2;
        if (coreWebView2 == null) return;
        
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(html));
        coreWebView2.ExecuteScriptAsync($"window.renderHtml(atob('{b64}'));void(0);");
    }
    catch (InvalidCastException)
    {
        return;
    }
    catch (Exception)
    {
        return;
    }
}
```

#### 3. Вставка текста вокруг выделения

```239:270:Controls/TaskPaneControl.cs
public void InsertInline(string prefix, string suffix)
{
    if (!_coreReady || _webView == null) return;
    try
    {
        var coreWebView2 = _webView.CoreWebView2;
        if (coreWebView2 == null) return;
        
        var p = Convert.ToBase64String(Encoding.UTF8.GetBytes(prefix ?? string.Empty));
        var s = Convert.ToBase64String(Encoding.UTF8.GetBytes(suffix ?? string.Empty));
        coreWebView2.ExecuteScriptAsync($"window.insertAroundSelection(atob('{p}'), atob('{s}'));void(0);");
    }
    catch (InvalidCastException)
    {
        return;
    }
    catch (Exception)
    {
        return;
    }
}
```

#### 4. Получение значения из JavaScript

```191:219:Controls/TaskPaneControl.cs
public async Task<string> GetMarkdownAsync()
{
    if (!string.IsNullOrEmpty(_latestMarkdown)) return _latestMarkdown;
    if (_coreReady && _webView != null)
    {
        try
        {
            var coreWebView2 = _webView.CoreWebView2;
            if (coreWebView2 != null)
            {
                var js = await coreWebView2.ExecuteScriptAsync("window.editorGetValue()");
                return UnquoteJsonString(js);
            }
        }
        catch (InvalidCastException)
        {
            return string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }
    return string.Empty;
}
```

**Особенность:** JavaScript возвращает JSON-строку (в кавычках), поэтому используется `UnquoteJsonString()` для обработки.

---

## Взаимодействие JavaScript → C#

### Механизм отправки

JavaScript отправляет сообщения через `window.chrome.webview.postMessage()`:

```javascript
// В JavaScript (из HTML)
function postToHost(type, text) {
    const b64 = btoa(encodeURIComponent(text || ''));
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(type + '|' + b64);
    }
}
```

### Обработка в C#

```70:106:Controls/TaskPaneControl.cs
private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e) 
{
    try 
    { 
        var json = e.TryGetWebMessageAsString();
        if (string.IsNullOrEmpty(json)) return;
        var parts = json.Split(new[] { '|' }, 2);
        if (parts.Length != 2) return;
        var type = parts[0];
        var payload = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1]));
        
        if (type == "mdChanged")
        {
            _latestMarkdown = payload;
            var html = _renderer.RenderoHtml(payload);
            PostRenderHtml(html);
        }
        
        if (type == "viewModeChanged")
        {
            try
            {
                Settings.Default.ViewMode = payload;
                Settings.Default.Save();
            }
            catch { }
        }
    }
    catch { }
}
```

### Типы сообщений

1. **`mdChanged`** - изменение Markdown текста
   - JavaScript отправляет при каждом изменении в редакторе (с debounce 120ms)
   - C# получает, сохраняет в кэш, конвертирует в HTML и отправляет обратно для preview

2. **`viewModeChanged`** - изменение режима отображения
   - JavaScript отправляет при переключении режима (split/markdown/html)
   - C# сохраняет в настройки приложения

---

## Формат передачи данных

### Протокол обмена

Все данные передаются в формате: `"тип|данныеВBase64"`

### Примеры преобразования

#### Пример 1: Отправка Markdown из JS в C#

```
Исходный текст: "Hello **World**"
↓
JavaScript: btoa(encodeURIComponent("Hello **World**"))
↓
Base64: "SGVsbG8gKipXb3JsZCoq"
↓
Сообщение: "mdChanged|SGVsbG8gKipXb3JsZCoq"
↓
C#: Split по '|' → ["mdChanged", "SGVsbG8gKipXb3JsZCoq"]
↓
C#: Convert.FromBase64String() → байты
↓
C#: Encoding.UTF8.GetString() → "Hello **World**"
```

#### Пример 2: Отправка HTML из C# в JS

```
Исходный HTML: "<p>Hello</p>"
↓
C#: Encoding.UTF8.GetBytes() → [60, 112, 62, 72, 101, 108, 108, 111, 60, 47, 112, 62]
↓
C#: Convert.ToBase64String() → "PHA+SGVsbG8gKipXb3JsZCoq"
↓
JS вызов: window.renderHtml(atob('PHA+SGVsbG8gKipXb3JsZCoq'))
↓
JavaScript: atob() → "<p>Hello</p>"
↓
Отображается в preview
```

### Почему Base64?

1. **Безопасность**: Избегает проблем с экранированием специальных символов
2. **Надежность**: Гарантирует корректную передачу любых UTF-8 символов
3. **Совместимость**: Стандартный формат для передачи бинарных данных через текстовые протоколы

---

## Потоки данных

### Поток 1: Пользователь вводит текст в редактор

```
1. Пользователь вводит: "Hello"
   ↓
2. JavaScript: Событие 'input' на textarea
   ↓
3. JavaScript: debounce(120ms) → notifyChange()
   ↓
4. JavaScript: postToHost('mdChanged', 'Hello')
   ↓
5. JavaScript: Кодирование в Base64 → "mdChanged|SGVsbG8="
   ↓
6. WebView2: postMessage → C#
   ↓
7. C#: CoreWebView2_WebMessageReceived
   ↓
8. C#: Декодирование → "Hello"
   ↓
9. C#: Сохранение в _latestMarkdown
   ↓
10. C#: _renderer.RenderoHtml("Hello") → "<p>Hello</p>"
   ↓
11. C#: PostRenderHtml("<p>Hello</p>")
   ↓
12. C#: Кодирование в Base64 → ExecuteScriptAsync("window.renderHtml(...)")
   ↓
13. JavaScript: window.renderHtml() → декодирование → отображение в preview
```

### Поток 2: C# загружает файл в редактор

```
1. C#: OpenMarkdownFile() → File.ReadAllText()
   ↓
2. C#: SetMarkdown(fileContent)
   ↓
3. C#: Кодирование в Base64
   ↓
4. C#: ExecuteScriptAsync("window.editorSetValue(atob('...'))")
   ↓
5. JavaScript: window.editorSetValue() → декодирование
   ↓
6. JavaScript: editor.value = decodedText
   ↓
7. JavaScript: notifyChange() → отправка обратно в C#
   ↓
8. C#: Получение и обработка (как в Потоке 1)
```

### Поток 3: Переключение режима отображения

```
1. Пользователь нажимает кнопку "Markdown"
   ↓
2. JavaScript: setViewMode('markdown')
   ↓
3. JavaScript: applyViewMode('markdown', true)
   ↓
4. JavaScript: Изменение CSS классов на body
   ↓
5. JavaScript: postToHost('viewModeChanged', 'markdown')
   ↓
6. C#: CoreWebView2_WebMessageReceived → type = "viewModeChanged"
   ↓
7. C#: Settings.Default.ViewMode = "markdown"
   ↓
8. C#: Settings.Default.Save()
```

---

## Основные функции и методы

### JavaScript функции (доступны из C#)

#### `window.editorSetValue(text)`
Устанавливает значение в textarea редактора.

```javascript
window.editorSetValue = function(text) { 
    editor.value = text || ''; 
    notifyChange(); 
}
```

#### `window.editorGetValue()`
Возвращает текущее значение редактора.

```javascript
window.editorGetValue = function() { 
    return editor.value || ''; 
}
```

#### `window.renderHtml(html)`
Отображает HTML в preview с обработкой:
- Санитизация через DOMPurify
- Подсветка кода через Prism.js
- Рендеринг Mermaid диаграмм
- Рендеринг MathJax формул

```javascript
window.renderHtml = function(html) {
    const clean = DOMPurify.sanitize(html || '', { 
        ADD_ATTR: ['target', 'rel', 'class', 'style', 'id'] 
    });
    preview.innerHTML = clean;
    // ... обработка Mermaid, Prism, MathJax
}
```

#### `window.insertAroundSelection(prefix, suffix)`
Вставляет текст вокруг выделенного фрагмента.

#### `window.insertSnippet(snippet)`
Вставляет текст в позицию курсора.

#### `window.setViewMode(mode)`
Переключает режим отображения (split/markdown/html).

#### `window.getViewMode()`
Возвращает текущий режим отображения.

### C# методы

#### Публичные методы

- `SetMarkdown(string)` - установка Markdown текста
- `GetMarkdownAsync()` - получение Markdown текста
- `GetCachedMarkdown()` - получение из кэша (синхронно)
- `InsertInline(prefix, suffix)` - вставка вокруг выделения
- `InsertSnippet(snippet)` - вставка сниппета
- `InsertHeading(level)` - вставка заголовка
- `InsertBulletList()` - вставка маркированного списка
- `InsertNumberedList()` - вставка нумерованного списка
- `InsertCheckbox(isChecked)` - вставка чекбокса
- `InsertTable(rows, cols)` - вставка таблицы
- `InsertLink(text, url)` - вставка ссылки
- `InsertImage(alt, path)` - вставка изображения
- `InsertCodeBlock(language)` - вставка блока кода
- `InsertMermaid(text)` - вставка Mermaid диаграммы
- `InsertMath(text)` - вставка математической формулы
- `SaveMarkdownFile()` - сохранение в файл
- `OpenMarkdownFile()` - открытие файла
- `SetViewMode(mode)` - установка режима отображения
- `GetCurrentViewModeAsync()` - получение текущего режима

#### Приватные методы

- `OnLoadAsync()` - инициализация WebView2
- `CoreWebView2_WebMessageReceived()` - обработка сообщений из JS
- `PostRenderHtml(html)` - отправка HTML в preview
- `UnquoteJsonString(json)` - обработка JSON-строк из JS
- `BuildHtmlShell()` - генерация HTML документа

---

## Примеры взаимодействия

### Пример 1: Пользователь вводит заголовок

**Сценарий:** Пользователь вводит `# Заголовок` в редактор.

```
1. [HTML] textarea: пользователь вводит "# Заголовок"
   ↓
2. [JS] Событие 'input' → debounce(120ms)
   ↓
3. [JS] notifyChange() → postToHost('mdChanged', '# Заголовок')
   ↓
4. [JS] Кодирование: "mdChanged|IyDQl9Cw0LPQu9C10LbQvtC8"
   ↓
5. [WebView2] postMessage → C#
   ↓
6. [C#] CoreWebView2_WebMessageReceived
   ↓
7. [C#] Split: ["mdChanged", "IyDQl9Cw0LPQu9C10LbQvtC8"]
   ↓
8. [C#] Декодирование: "# Заголовок"
   ↓
9. [C#] _latestMarkdown = "# Заголовок"
   ↓
10. [C#] _renderer.RenderoHtml("# Заголовок") → "<h1>Заголовок</h1>"
   ↓
11. [C#] PostRenderHtml("<h1>Заголовок</h1>")
   ↓
12. [C#] Кодирование: "PHA+0JfQsNCz0LvQtdC20L7QvDwvcD4="
   ↓
13. [C#] ExecuteScriptAsync("window.renderHtml(atob('PHA+...'))")
   ↓
14. [JS] window.renderHtml() → декодирование → "<h1>Заголовок</h1>"
   ↓
15. [JS] preview.innerHTML = "<h1>Заголовок</h1>"
   ↓
16. [HTML] Отображение заголовка в preview
```

### Пример 2: C# вставляет жирный текст

**Сценарий:** Вызов `InsertInline("**", "**")` для выделенного текста "Hello".

```
1. [C#] InsertInline("**", "**")
   ↓
2. [C#] Кодирование prefix: "**" → "Kg=="
   ↓
3. [C#] Кодирование suffix: "**" → "Kg=="
   ↓
4. [C#] ExecuteScriptAsync("window.insertAroundSelection(atob('Kg=='), atob('Kg=='))")
   ↓
5. [JS] window.insertAroundSelection() → декодирование → ("**", "**")
   ↓
6. [JS] editor.value = "**Hello**" (вокруг выделения)
   ↓
7. [JS] notifyChange() → отправка в C#
   ↓
8. [C#] Получение "**Hello**" → конвертация → отображение в preview
```

### Пример 3: Загрузка файла

**Сценарий:** Пользователь открывает файл `document.md`.

```
1. [C#] OpenMarkdownFile() → OpenFileDialog
   ↓
2. [C#] File.ReadAllText("document.md") → "## Заголовок\n\nТекст"
   ↓
3. [C#] SetMarkdown("## Заголовок\n\nТекст")
   ↓
4. [C#] Кодирование в Base64
   ↓
5. [C#] ExecuteScriptAsync("window.editorSetValue(atob('...'))")
   ↓
6. [JS] window.editorSetValue() → editor.value = "## Заголовок\n\nТекст"
   ↓
7. [JS] notifyChange() → отправка обратно в C#
   ↓
8. [C#] Обработка → конвертация → preview
   ↓
9. [C#] DocumentSyncService.SaveMarkdownToActiveDocument() → синхронизация с Word
```

---

## Кэширование данных

### Механизм кэширования

C# поддерживает локальный кэш `_latestMarkdown` для быстрого доступа:

```18:19:Controls/TaskPaneControl.cs
private string _latestMarkdown = string.Empty;
private bool _coreReady = false;
```

**Преимущества:**
- Быстрый доступ без запроса к JavaScript
- Работает даже если WebView2 не готов
- Синхронизируется автоматически через события `mdChanged`

**Использование:**
- `GetCachedMarkdown()` - мгновенный возврат из кэша
- `GetMarkdownAsync()` - сначала проверяет кэш, затем запрашивает из JS

---

## Обработка ошибок

### Защитные механизмы

1. **Проверка готовности WebView2:**
   ```csharp
   if (!_coreReady || _webView == null) return;
   ```

2. **Обработка исключений:**
   ```csharp
   catch (InvalidCastException) { return; }  // WebView2 не готов
   catch (Exception) { return; }             // Другие ошибки
   ```

3. **Валидация данных:**
   ```csharp
   if (parts.Length != 2) return;  // Неверный формат сообщения
   ```

4. **Защита от null:**
   ```csharp
   var coreWebView2 = _webView.CoreWebView2;
   if (coreWebView2 == null) return;
   ```

---

## Заключение

Модуль `TaskPaneControl` реализует двустороннюю коммуникацию между C# и JavaScript через WebView2:

- **C# → JS**: Вызов функций через `ExecuteScriptAsync()`
- **JS → C#**: Отправка сообщений через `postMessage()`
- **Формат**: Все данные передаются в Base64 для безопасности
- **Кэширование**: Локальный кэш для быстрого доступа
- **Обработка ошибок**: Множественные уровни защиты

Система обеспечивает надежную работу Markdown-редактора с предпросмотром в реальном времени, интеграцией с Word и поддержкой расширенных возможностей (Mermaid, MathJax, Prism.js).

