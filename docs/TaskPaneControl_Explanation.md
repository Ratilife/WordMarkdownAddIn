# Объяснение работы класса TaskPaneControl

## Обзор архитектуры

Класс `TaskPaneControl` является мостом между **C#** (бэкенд Word Add-In) и **JavaScript/HTML** (веб-интерфейс редактора). Он использует **WebView2** для отображения HTML-страницы с JavaScript-редактором Markdown.

---

## Схема взаимодействия языков

```
┌─────────────────────────────────────────────────────────────────┐
│                         C# (TaskPaneControl.cs)                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  UserControl (Windows Forms)                             │   │
│  │  ┌────────────────────────────────────────────────────┐  │   │
│  │  │  WebView2 (Microsoft Edge WebView2)                │  │   │
│  │  │  ┌──────────────────────────────────────────────┐   │  │   │
│  │  │  │  HTML (BuildHtmlShell)                      │   │  │   │
│  │  │  │  ┌────────────────────────────────────────┐  │   │  │   │
│  │  │  │  │  JavaScript (редактор + обработчики)  │  │   │  │   │
│  │  │  │  └────────────────────────────────────────┘  │   │  │   │
│  │  │  └──────────────────────────────────────────────┘   │  │   │
│  │  └────────────────────────────────────────────────────┘  │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

### Детальная схема потока данных

```
┌─────────────────────────────────────────────────────────────────────┐
│                          ПОТОК ДАННЫХ                                │
└─────────────────────────────────────────────────────────────────────┘

┌──────────────┐                    ┌──────────────┐
│     C#      │                    │ JavaScript   │
│             │                    │              │
│  TaskPane   │◄───Base64───────►│  HTML Editor  │
│  Control    │   Messages         │              │
│             │                    │              │
│  ┌────────┐│                    │  ┌─────────┐ │
│  │Markdown││                    │  │<textarea>│ │
│  │Renderer││                    │  │#editor  │ │
│  └────────┘│                    │  └─────────┘ │
│     │      │                    │       │       │
│     │      │                    │       │       │
│     ▼      │                    │       ▼       │
│  ┌─────┐   │                    │  ┌───────┐   │
│  │HTML │   │                    │  │Preview│   │
│  └─────┘   │                    │  └───────┘   │
└──────────────┘                    └──────────────┘
       │                                    │
       │                                    │
       └──────────────┬─────────────────────┘
                      │
                      ▼
            ┌─────────────────┐
            │  WebView2 API   │
            │  (Bridge)       │
            └─────────────────┘
```

---

## Последовательность вызовов кода

### 1. Инициализация (при загрузке контрола)

```
┌─────────────────────────────────────────────────────────────┐
│ ШАГ 1: Конструктор TaskPaneControl()                         │
│                                                              │
│  1.1. Создается экземпляр MarkdownRenderService             │
│  1.2. Создается WebView2 компонент                           │
│  1.3. WebView2 добавляется в Controls (UserControl)          │
│  1.4. Подписывается обработчик Load → OnLoadAsync           │
└─────────────────────────────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│ ШАГ 2: OnLoadAsync() - Асинхронная инициализация            │
│                                                              │
│  2.1. await _webView.EnsureCoreWebView2Async()              │
│       └─► Инициализирует WebView2 Runtime                   │
│       └─► Создает браузерный процесс                         │
│                                                              │
│  2.2. _coreReady = true                                      │
│       └─► Разрешает выполнение методов                       │
│                                                              │
│  2.3. Подписка на WebMessageReceived                         │
│       └─► CoreWebView2_WebMessageReceived                   │
│                                                              │
│  2.4. Настройка параметров WebView2:                        │
│       - DevTools включены                                    │
│       - StatusBar выключен                                   │
│       - ContextMenus включены                                │
│                                                              │
│  2.5. _webView.CoreWebView2.NavigateToString(               │
│           BuildHtmlShell()                                   │
│       )                                                      │
│       └─► Загружает HTML в WebView2                         │
└─────────────────────────────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│ ШАГ 3: BuildHtmlShell() - Генерация HTML                    │
│                                                              │
│  3.1. Возвращает полный HTML документ:                       │
│       - <head> с CSS стилями                                 │
│       - <body> с:                                            │
│         • Панель управления режимами (Split/Markdown/HTML)  │
│         • <textarea id="editor"> - редактор Markdown         │
│         • <div id="preview"> - предпросмотр HTML             │
│       - <script> блоки:                                      │
│         • DOMPurify (санитизация)                            │
│         • Prism.js (подсветка синтаксиса)                    │
│         • Mermaid (диаграммы)                                │
│         • MathJax (математические формулы)                   │
│         • JavaScript код редактора                           │
└─────────────────────────────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│ ШАГ 4: JavaScript инициализация (в HTML)                    │
│                                                              │
│  4.1. Загружаются внешние библиотеки (CDN)                  │
│                                                              │
│  4.2. Инициализируются переменные:                           │
│       - const editor = document.getElementById('editor')     │
│       - const preview = document.getElementById('preview')  │
│                                                              │
│  4.3. Регистрируются обработчики событий:                    │
│       - editor.addEventListener('input', debounce(...))      │
│       - Кнопки переключения режимов                          │
│                                                              │
│  4.4. Определяются функции для C#:                          │
│       - window.editorSetValue()                              │
│       - window.editorGetValue()                              │
│       - window.insertAroundSelection()                       │
│       - window.insertSnippet()                              │
│       - window.renderHtml()                                  │
│       - window.setViewMode()                                 │
│       - window.getViewMode()                                 │
│                                                              │
│  4.5. setTimeout() - инициализация после загрузки:           │
│       - Загрузка сохраненного режима                         │
│       - editor.focus()                                       │
│       - notifyChange() - отправка начального состояния       │
└─────────────────────────────────────────────────────────────┘
```

### 2. Типичный цикл работы (пользователь редактирует текст)

```
┌─────────────────────────────────────────────────────────────┐
│ ЦИКЛ: Редактирование Markdown                                │
└─────────────────────────────────────────────────────────────┘

┌──────────────┐
│  Пользователь│
│  вводит текст│
└──────┬───────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ JavaScript: editor.addEventListener('input', ...)            │
│                                                              │
│  1. Срабатывает событие 'input' на textarea                  │
│  2. debounce() задерживает вызов на 120ms                   │
│  3. Вызывается notifyChange()                                │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ JavaScript: notifyChange()                                   │
│                                                              │
│  postToHost('mdChanged', editor.value)                      │
│    │                                                          │
│    ├─► btoa(encodeURIComponent(text))                       │
│    │   └─► Кодирование в Base64                             │
│    │                                                          │
│    └─► window.chrome.webview.postMessage(                   │
│          'mdChanged|SGVsbG8gV29ybGQ='                        │
│        )                                                      │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ C#: CoreWebView2_WebMessageReceived()                        │
│                                                              │
│  1. Получение сообщения: "mdChanged|SGVsbG8gV29ybGQ="       │
│                                                              │
│  2. Разделение по '|':                                       │
│     - type = "mdChanged"                                     │
│     - payload_b64 = "SGVsbG8gV29ybGQ="                       │
│                                                              │
│  3. Декодирование Base64:                                    │
│     Convert.FromBase64String() → байты                      │
│     Encoding.UTF8.GetString() → "Hello World"                │
│                                                              │
│  4. Сохранение в кэш:                                        │
│     _latestMarkdown = "Hello World"                          │
│                                                              │
│  5. Конвертация Markdown → HTML:                             │
│     _renderer.RenderoHtml("Hello World")                     │
│     └─► Возвращает: "<p>Hello World</p>"                     │
│                                                              │
│  6. Вызов PostRenderHtml("<p>Hello World</p>")              │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ C#: PostRenderHtml()                                         │
│                                                              │
│  1. Кодирование HTML в Base64:                              │
│     Encoding.UTF8.GetBytes("<p>Hello World</p>")            │
│     Convert.ToBase64String() → "PHA+SGVsbG8gV29ybGQ8L3A+"    │
│                                                              │
│  2. Выполнение JavaScript:                                   │
│     ExecuteScriptAsync(                                      │
│       "window.renderHtml(atob('PHA+SGVsbG8gV29ybGQ8L3A+'));"│
│     )                                                        │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ JavaScript: window.renderHtml()                             │
│                                                              │
│  1. Декодирование: atob('PHA+SGVsbG8gV29ybGQ8L3A+')         │
│     └─► "<p>Hello World</p>"                                 │
│                                                              │
│  2. Санитизация: DOMPurify.sanitize(html)                   │
│                                                              │
│  3. Отображение: preview.innerHTML = clean                   │
│                                                              │
│  4. Обработка специальных элементов:                        │
│     - Prism.highlightAllUnder() - подсветка кода            │
│     - mermaid.init() - рендеринг диаграмм                    │
│     - MathJax.typesetPromise() - рендеринг формул           │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌──────────────┐
│  Пользователь│
│  видит HTML  │
│  в preview   │
└──────────────┘
```

### 3. Вызов методов из C# (например, вставка заголовка)

```
┌─────────────────────────────────────────────────────────────┐
│ ВЫЗОВ: InsertHeading(2) из Ribbon или другого места         │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│ C#: InsertHeading(2)                                        │
│                                                              │
│  1. Валидация: level = 2 (в пределах 1-6)                   │
│                                                              │
│  2. Генерация Markdown:                                      │
│     new string('#', 2) → "##"                                │
│     "\n## "                                                  │
│                                                              │
│  3. Вызов InsertSnippet("\n## ")                            │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ C#: InsertSnippet("\n## ")                                  │
│                                                              │
│  1. Кодирование в Base64:                                    │
│     Convert.ToBase64String(                                  │
│       Encoding.UTF8.GetBytes("\n## ")                       │
│     ) → "CgojIyA="                                           │
│                                                              │
│  2. Выполнение JavaScript:                                   │
│     ExecuteScriptAsync(                                      │
│       "window.insertSnippet(atob('CgojIyA='));"             │
│     )                                                        │
└─────────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────┐
│ JavaScript: window.insertSnippet()                          │
│                                                              │
│  1. Декодирование: atob('CgojIyA=') → "\n## "               │
│                                                              │
│  2. Получение позиции курсора:                              │
│     const pos = editor.selectionStart                        │
│                                                              │
│  3. Вставка текста:                                          │
│     editor.value = val.substring(0, pos) +                   │
│                    "\n## " +                                │
│                    val.substring(pos)                        │
│                                                              │
│  4. Установка курсора после вставки:                        │
│     editor.setSelectionRange(newPos, newPos)                │
│                                                              │
│  5. Фокус на редактор: editor.focus()                       │
│                                                              │
│  6. Уведомление об изменении: notifyChange()                │
│     └─► Запускается цикл обработки (см. выше)                │
└─────────────────────────────────────────────────────────────┘
```

---

## Детальное объяснение взаимодействия

### 1. C# → JavaScript (Вызов функций)

**Механизм:** `ExecuteScriptAsync()`

**Процесс:**
1. C# кодирует данные в Base64
2. C# формирует строку JavaScript кода
3. WebView2 выполняет JavaScript код
4. JavaScript декодирует Base64 и выполняет действие

**Пример:**
```csharp
// C# код
var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes("Hello"));
_webView.CoreWebView2.ExecuteScriptAsync($"window.editorSetValue(atob('{b64}'));");
```

```javascript
// JavaScript код (выполняется в WebView2)
window.editorSetValue = function(text) {
    editor.value = text || '';  // text = "Hello" (уже декодирован atob)
    notifyChange();
}
```

### 2. JavaScript → C# (Отправка сообщений)

**Механизм:** `WebMessageReceived` событие

**Процесс:**
1. JavaScript кодирует данные в Base64
2. JavaScript отправляет сообщение через `postMessage()`
3. WebView2 передает сообщение в C#
4. C# получает событие `WebMessageReceived`
5. C# декодирует Base64 и обрабатывает данные

**Пример:**
```javascript
// JavaScript код
function postToHost(type, text) {
    const b64 = btoa(encodeURIComponent(text || ''));
    window.chrome.webview.postMessage(type + '|' + b64);
    // Отправляет: "mdChanged|SGVsbG8="
}
```

```csharp
// C# код
private void CoreWebView2_WebMessageReceived(object sender, ...) {
    var json = e.TryGetWebMessageAsString();  // "mdChanged|SGVsbG8="
    var parts = json.Split('|');               // ["mdChanged", "SGVsbG8="]
    var payload = Encoding.UTF8.GetString(
        Convert.FromBase64String(parts[1])    // "Hello"
    );
    // Обработка...
}
```

### 3. Почему Base64?

**Причины использования Base64:**
- **Безопасность:** Избегает проблем с экранированием специальных символов
- **Надежность:** Гарантирует корректную передачу UTF-8 символов
- **Совместимость:** Работает через строковый интерфейс WebView2

**Пример проблемы без Base64:**
```csharp
// ❌ ПРОБЛЕМА: Специальные символы могут сломать JavaScript
ExecuteScriptAsync($"window.setValue('Hello\nWorld');");
// Ошибка: неожиданный символ новой строки
```

```csharp
// ✅ РЕШЕНИЕ: Base64 кодирование
var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes("Hello\nWorld"));
ExecuteScriptAsync($"window.setValue(atob('{b64}'));");
// Работает корректно
```

---

## Ключевые компоненты

### 1. WebView2
- **Что это:** Компонент Microsoft для встраивания веб-контента в Windows приложения
- **Роль:** Мост между C# и JavaScript
- **API:**
  - `NavigateToString()` - загрузка HTML
  - `ExecuteScriptAsync()` - выполнение JavaScript
  - `WebMessageReceived` - получение сообщений от JavaScript

### 2. HTML Shell
- **Что это:** Полный HTML документ, генерируемый методом `BuildHtmlShell()`
- **Содержит:**
  - CSS стили для редактора и предпросмотра
  - HTML структуру (textarea, div для preview)
  - JavaScript код для логики редактора
  - Подключение внешних библиотек (CDN)

### 3. JavaScript API
- **Функции для C#:**
  - `window.editorSetValue(text)` - установка текста
  - `window.editorGetValue()` - получение текста
  - `window.insertAroundSelection(prefix, suffix)` - вставка вокруг выделения
  - `window.insertSnippet(snippet)` - вставка сниппета
  - `window.renderHtml(html)` - отображение HTML в preview
  - `window.setViewMode(mode)` - переключение режима
  - `window.getViewMode()` - получение режима

### 4. MarkdownRenderService
- **Что это:** C# сервис для конвертации Markdown в HTML
- **Использует:** Markdig библиотеку
- **Роль:** Преобразует Markdown текст в HTML для отображения в preview

---

## Схема архитектуры компонентов

```
┌──────────────────────────────────────────────────────────────┐
│                    Word Add-In Application                    │
│                                                               │
│  ┌────────────────────────────────────────────────────────┐ │
│  │  ThisAddIn.cs                                           │ │
│  │  └─► Создает TaskPaneControl                            │ │
│  └────────────────────────────────────────────────────────┘ │
│                        │                                      │
│                        ▼                                      │
│  ┌────────────────────────────────────────────────────────┐ │
│  │  TaskPaneControl.cs (C#)                                │ │
│  │  ┌──────────────────────────────────────────────────┐  │ │
│  │  │  WebView2                                        │  │ │
│  │  │  ┌────────────────────────────────────────────┐  │  │ │
│  │  │  │  HTML (BuildHtmlShell)                    │  │  │ │
│  │  │  │  ┌──────────────────────────────────────┐│  │  │ │
│  │  │  │  │  JavaScript Editor                    ││  │  │ │
│  │  │  │  │  - textarea#editor                    ││  │  │ │
│  │  │  │  │  - div#preview                       ││  │  │ │
│  │  │  │  │  - postToHost()                      ││  │  │ │
│  │  │  │  │  - window.* функции                   ││  │  │ │
│  │  │  │  └──────────────────────────────────────┘│  │  │ │
│  │  │  └────────────────────────────────────────────┘  │  │ │
│  │  └──────────────────────────────────────────────────┘  │  │
│  │  ┌──────────────────────────────────────────────────┐  │  │
│  │  │  MarkdownRenderService                            │  │  │
│  │  │  └─► Markdown → HTML                              │  │  │
│  │  └──────────────────────────────────────────────────┘  │  │
│  └────────────────────────────────────────────────────────┘ │
│                        │                                      │
│                        ▼                                      │
│  ┌────────────────────────────────────────────────────────┐ │
│  │  Ribbon (MarkdownRibbon.cs)                            │ │
│  │  └─► Кнопки → вызовы методов TaskPaneControl            │ │
│  └────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────┘
```

---

## Резюме

### Как работает взаимодействие:

1. **C# управляет жизненным циклом:**
   - Создает WebView2
   - Загружает HTML
   - Обрабатывает сообщения от JavaScript

2. **HTML предоставляет интерфейс:**
   - Структура редактора
   - Стили оформления
   - Контейнер для JavaScript

3. **JavaScript реализует логику:**
   - Обработка ввода пользователя
   - Управление редактором
   - Отправка сообщений в C#

4. **WebView2 является мостом:**
   - `ExecuteScriptAsync()` - C# → JavaScript
   - `WebMessageReceived` - JavaScript → C#

### Формат передачи данных:

- **Всегда Base64** для безопасной передачи
- **Формат:** `"тип|данныеВBase64"`
- **Пример:** `"mdChanged|SGVsbG8gV29ybGQ="`

### Последовательность типичного действия:

1. Пользователь вводит текст → JavaScript событие
2. JavaScript кодирует и отправляет → C# получает
3. C# обрабатывает (рендерит Markdown) → отправляет HTML обратно
4. JavaScript получает HTML → отображает в preview

















