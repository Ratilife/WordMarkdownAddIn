# Полная инструкция по реализации модуля WordMarkdownFormatter

## Содержание
1. [Обзор реализации](#обзор-реализации)
2. [Этап 1: Создание базовой структуры классов](#этап-1-создание-базовой-структуры-классов)
3. [Этап 2: Реализация MarkdownPatternMatcher](#этап-2-реализация-markdownpatternmatcher)
4. [Этап 3: Реализация MarkdownElementFormatter](#этап-3-реализация-markdownelementformatter)
5. [Этап 4: Реализация WordMarkdownFormatter](#этап-4-реализация-wordmarkdownformatter)
6. [Этап 5: Интеграция с Ribbon](#этап-5-интеграция-с-ribbon)
7. [Этап 6: Тестирование](#этап-6-тестирование)
8. [Приложение: Полный код классов](#приложение-полный-код-классов)

---

## Обзор реализации

Модуль `WordMarkdownFormatter` будет состоять из следующих компонентов:

1. **WordMarkdownFormatter.cs** - основной класс-координатор
2. **MarkdownPatternMatcher** - класс для поиска Markdown-синтаксиса (вложенный или отдельный)
3. **MarkdownElementFormatter** - класс для применения форматирования (вложенный или отдельный)
4. **Вспомогательные классы**: `MarkdownElementMatch`, `MarkdownElementType`, `FormattingContext`

### Порядок реализации
Рекомендуется реализовывать модуль поэтапно, начиная с базовых структур данных, затем паттерн-матчер, форматтер, и наконец основной класс-координатор.

---

## Этап 1: Создание базовой структуры классов

### Шаг 1.1: Создание файла WordMarkdownFormatter.cs

**Действие:** Создайте новый файл `Services/WordMarkdownFormatter.cs`

**Содержимое файла (начальная структура):**

```csharp
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace WordMarkdownAddIn.Services
{
    /// <summary>
    /// Основной класс для форматирования Markdown-синтаксиса в документе Word.
    /// Находит элементы Markdown, применяет форматирование Word и удаляет синтаксические маркеры.
    /// </summary>
    public class WordMarkdownFormatter
    {
        // Приватные поля будут добавлены на следующих шагах
        
        /// <summary>
        /// Конструктор класса WordMarkdownFormatter
        /// </summary>
        public WordMarkdownFormatter()
        {
            // Инициализация будет добавлена позже
        }
    }
}
```

### Шаг 1.2: Создание перечисления MarkdownElementType

**Действие:** Добавьте перечисление в начало файла `WordMarkdownFormatter.cs` (перед классом)

```csharp
/// <summary>
/// Типы элементов Markdown, которые могут быть найдены в документе
/// </summary>
public enum MarkdownElementType
{
    Heading,        // Заголовок (# Заголовок)
    Bold,           // Жирный текст (**текст**)
    Italic,         // Курсив (*текст*)
    Strikethrough,  // Зачеркнутый текст (~~текст~~)
    InlineCode,     // Инлайн-код (`код`)
    CodeBlock,      // Блок кода (```код```)
    Link,           // Ссылка ([текст](url))
    ListItem,       // Элемент списка (- элемент, 1. элемент)
    Table,          // Таблица
    Quote,          // Цитата (> цитата)
    HorizontalRule  // Горизонтальная линия (---)
}
```

### Шаг 1.3: Создание класса MarkdownElementMatch

**Действие:** Добавьте класс `MarkdownElementMatch` в файл `WordMarkdownFormatter.cs` (перед основным классом)

```csharp
/// <summary>
/// Структура данных для хранения информации о найденном элементе Markdown
/// </summary>
public class MarkdownElementMatch
{
    /// <summary>
    /// Тип элемента Markdown
    /// </summary>
    public MarkdownElementType ElementType { get; set; }

    /// <summary>
    /// Позиция начала элемента в документе (в символах от начала)
    /// </summary>
    public int StartPosition { get; set; }

    /// <summary>
    /// Позиция конца элемента в документе (в символах от начала)
    /// </summary>
    public int EndPosition { get; set; }

    /// <summary>
    /// Извлеченное содержимое элемента (без синтаксиса Markdown)
    /// </summary>
    public string Content { get; set; } = "";

    /// <summary>
    /// Полное совпадение с синтаксисом Markdown (включая маркеры)
    /// </summary>
    public string FullMatch { get; set; } = "";

    /// <summary>
    /// Дополнительные метаданные элемента
    /// Например: уровень заголовка, URL ссылки, язык кода и т.д.
    /// </summary>
    public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();

    /// <summary>
    /// Длина элемента в символах
    /// </summary>
    public int Length => EndPosition - StartPosition;
}
```

### Шаг 1.4: Создание класса FormattingContext

**Действие:** Добавьте класс `FormattingContext` в файл `WordMarkdownFormatter.cs` (перед основным классом)

```csharp
/// <summary>
/// Вспомогательный класс для хранения контекста форматирования во время обработки
/// </summary>
private class FormattingContext
{
    /// <summary>
    /// Диапазон текста, который обрабатывается
    /// </summary>
    public Range TargetRange { get; set; }

    /// <summary>
    /// Текст диапазона (кэшированный для быстрого доступа)
    /// </summary>
    public string Text { get; set; } = "";

    /// <summary>
    /// Список всех найденных элементов Markdown
    /// </summary>
    public List<MarkdownElementMatch> FoundElements { get; set; } = new List<MarkdownElementMatch>();

    /// <summary>
    /// Смещение позиций при удалении синтаксиса (для корректного обновления позиций)
    /// </summary>
    public int PositionOffset { get; set; } = 0;
}
```

**Проверка:** Убедитесь, что проект компилируется без ошибок.

---

## Этап 2: Реализация MarkdownPatternMatcher

### Шаг 2.1: Создание класса MarkdownPatternMatcher

**Действие:** Добавьте класс `MarkdownPatternMatcher` в файл `WordMarkdownFormatter.cs` (как вложенный класс или отдельный класс в том же файле)

**Рекомендация:** Сделайте его вложенным приватным классом внутри `WordMarkdownFormatter` для инкапсуляции.

```csharp
/// <summary>
/// Класс для поиска и распознавания синтаксиса Markdown в тексте документа
/// </summary>
private class MarkdownPatternMatcher
{
    // Регулярные выражения для различных типов элементов Markdown
    private readonly Regex _headingPattern;
    private readonly Regex _boldPattern;
    private readonly Regex _italicPattern;
    private readonly Regex _strikethroughPattern;
    private readonly Regex _inlineCodePattern;
    private readonly Regex _linkPattern;
    private readonly Regex _listItemPattern;
    private readonly Regex _codeBlockPattern;
    private readonly Regex _tablePattern;
    private readonly Regex _quotePattern;
    private readonly Regex _horizontalRulePattern;

    /// <summary>
    /// Конструктор - инициализирует все регулярные выражения
    /// </summary>
    public MarkdownPatternMatcher()
    {
        // Заголовки: # Заголовок, ## Заголовок и т.д. (1-6 уровней)
        _headingPattern = new Regex(@"^(#{1,6})\s+(.+)$", RegexOptions.Multiline | RegexOptions.Compiled);

        // Жирный текст: **текст**
        _boldPattern = new Regex(@"\*\*(.+?)\*\*", RegexOptions.Compiled);

        // Курсив: *текст* (но не **текст**)
        _italicPattern = new Regex(@"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", RegexOptions.Compiled);

        // Зачеркнутый текст: ~~текст~~
        _strikethroughPattern = new Regex(@"~~(.+?)~~", RegexOptions.Compiled);

        // Инлайн-код: `код`
        _inlineCodePattern = new Regex(@"`([^`]+)`", RegexOptions.Compiled);

        // Ссылки: [текст](url)
        _linkPattern = new Regex(@"\[([^\]]+)\]\(([^\)]+)\)", RegexOptions.Compiled);

        // Элементы списка: - элемент, * элемент, + элемент или 1. элемент
        _listItemPattern = new Regex(@"^([\-\*\+]|\d+\.)\s+(.+)$", RegexOptions.Multiline | RegexOptions.Compiled);

        // Блоки кода: ```код``` или ```language\nкод\n```
        _codeBlockPattern = new Regex(@"```(\w+)?\n([\s\S]*?)```", RegexOptions.Compiled);

        // Цитаты: > цитата
        _quotePattern = new Regex(@"^>\s+(.+)$", RegexOptions.Multiline | RegexOptions.Compiled);

        // Горизонтальная линия: --- или ***
        _horizontalRulePattern = new Regex(@"^(\-{3,}|\*{3,}|_{3,})$", RegexOptions.Multiline | RegexOptions.Compiled);

        // Таблицы - более сложный паттерн (будет реализован отдельно)
        _tablePattern = new Regex(@"^\|(.+)\|$", RegexOptions.Multiline | RegexOptions.Compiled);
    }
}
```

### Шаг 2.2: Реализация метода FindHeadings

**Действие:** Добавьте метод `FindHeadings` в класс `MarkdownPatternMatcher`

```csharp
/// <summary>
/// Поиск всех заголовков в тексте
/// </summary>
/// <param name="text">Текст для поиска</param>
/// <returns>Список найденных заголовков</returns>
public List<MarkdownElementMatch> FindHeadings(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _headingPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 3)
        {
            string hashes = match.Groups[1].Value;
            string content = match.Groups[2].Value;
            int level = hashes.Length;

            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.Heading,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = content,
                FullMatch = match.Value,
                Metadata = new Dictionary<string, object>
                {
                    { "Level", level }
                }
            };

            matches.Add(element);
        }
    }

    return matches;
}
```

### Шаг 2.3: Реализация методов поиска форматирования текста

**Действие:** Добавьте методы для поиска жирного текста, курсива и зачеркнутого текста

```csharp
/// <summary>
/// Поиск жирного текста
/// </summary>
public List<MarkdownElementMatch> FindBoldText(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _boldPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 2)
        {
            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.Bold,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = match.Groups[1].Value,
                FullMatch = match.Value
            };

            matches.Add(element);
        }
    }

    return matches;
}

/// <summary>
/// Поиск курсива
/// </summary>
public List<MarkdownElementMatch> FindItalicText(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _italicPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 2)
        {
            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.Italic,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = match.Groups[1].Value,
                FullMatch = match.Value
            };

            matches.Add(element);
        }
    }

    return matches;
}

/// <summary>
/// Поиск зачеркнутого текста
/// </summary>
public List<MarkdownElementMatch> FindStrikethroughText(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _strikethroughPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 2)
        {
            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.Strikethrough,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = match.Groups[1].Value,
                FullMatch = match.Value
            };

            matches.Add(element);
        }
    }

    return matches;
}
```

### Шаг 2.4: Реализация методов поиска кода и ссылок

**Действие:** Добавьте методы для поиска инлайн-кода, блоков кода и ссылок

```csharp
/// <summary>
/// Поиск инлайн-кода
/// </summary>
public List<MarkdownElementMatch> FindInlineCode(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _inlineCodePattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 2)
        {
            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.InlineCode,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = match.Groups[1].Value,
                FullMatch = match.Value
            };

            matches.Add(element);
        }
    }

    return matches;
}

/// <summary>
/// Поиск блоков кода
/// </summary>
public List<MarkdownElementMatch> FindCodeBlocks(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _codeBlockPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        string language = match.Groups.Count >= 2 && !string.IsNullOrEmpty(match.Groups[1].Value) 
            ? match.Groups[1].Value 
            : "";
        string code = match.Groups.Count >= 3 ? match.Groups[2].Value : "";

        var element = new MarkdownElementMatch
        {
            ElementType = MarkdownElementType.CodeBlock,
            StartPosition = match.Index,
            EndPosition = match.Index + match.Length,
            Content = code,
            FullMatch = match.Value,
            Metadata = new Dictionary<string, object>
            {
                { "Language", language }
            }
        };

        matches.Add(element);
    }

    return matches;
}

/// <summary>
/// Поиск ссылок
/// </summary>
public List<MarkdownElementMatch> FindLinks(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _linkPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 3)
        {
            string linkText = match.Groups[1].Value;
            string url = match.Groups[2].Value;

            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.Link,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = linkText,
                FullMatch = match.Value,
                Metadata = new Dictionary<string, object>
                {
                    { "Url", url }
                }
            };

            matches.Add(element);
        }
    }

    return matches;
}
```

### Шаг 2.5: Реализация методов поиска списков, цитат и таблиц

**Действие:** Добавьте оставшиеся методы поиска

```csharp
/// <summary>
/// Поиск элементов списков
/// </summary>
public List<MarkdownElementMatch> FindListItems(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _listItemPattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 3)
        {
            string marker = match.Groups[1].Value;
            string content = match.Groups[2].Value;
            bool isOrdered = char.IsDigit(marker[0]);

            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.ListItem,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = content,
                FullMatch = match.Value,
                Metadata = new Dictionary<string, object>
                {
                    { "IsOrdered", isOrdered },
                    { "Marker", marker }
                }
            };

            matches.Add(element);
        }
    }

    return matches;
}

/// <summary>
/// Поиск цитат
/// </summary>
public List<MarkdownElementMatch> FindQuotes(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _quotePattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        if (match.Groups.Count >= 2)
        {
            var element = new MarkdownElementMatch
            {
                ElementType = MarkdownElementType.Quote,
                StartPosition = match.Index,
                EndPosition = match.Index + match.Length,
                Content = match.Groups[1].Value,
                FullMatch = match.Value
            };

            matches.Add(element);
        }
    }

    return matches;
}

/// <summary>
/// Поиск горизонтальных линий
/// </summary>
public List<MarkdownElementMatch> FindHorizontalRules(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    var regexMatches = _horizontalRulePattern.Matches(text);
    foreach (Match match in regexMatches)
    {
        var element = new MarkdownElementMatch
        {
            ElementType = MarkdownElementType.HorizontalRule,
            StartPosition = match.Index,
            EndPosition = match.Index + match.Length,
            Content = "",
            FullMatch = match.Value
        };

        matches.Add(element);
    }

    return matches;
}

/// <summary>
/// Поиск таблиц (упрощенная версия)
/// ВАЖНО: Полная реализация парсинга таблиц требует более сложной логики
/// </summary>
public List<MarkdownElementMatch> FindTables(string text)
{
    var matches = new List<MarkdownElementMatch>();

    if (string.IsNullOrEmpty(text))
        return matches;

    // Упрощенная реализация - поиск строк таблицы
    // Полная реализация должна парсить структуру таблицы (строки, столбцы)
    var lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
    int currentPosition = 0;
    List<string> tableLines = new List<string>();
    int tableStart = -1;

    for (int i = 0; i < lines.Length; i++)
    {
        string line = lines[i];
        if (_tablePattern.IsMatch(line.Trim()))
        {
            if (tableStart == -1)
            {
                tableStart = currentPosition;
            }
            tableLines.Add(line);
        }
        else
        {
            // Если накопили строки таблицы, создаем элемент
            if (tableLines.Count >= 2 && tableStart != -1)
            {
                string tableText = string.Join("\n", tableLines);
                var element = new MarkdownElementMatch
                {
                    ElementType = MarkdownElementType.Table,
                    StartPosition = tableStart,
                    EndPosition = tableStart + tableText.Length,
                    Content = tableText,
                    FullMatch = tableText,
                    Metadata = new Dictionary<string, object>
                    {
                        { "RowCount", tableLines.Count }
                    }
                };

                matches.Add(element);
            }

            tableLines.Clear();
            tableStart = -1;
        }

        currentPosition += line.Length + 1; // +1 для символа новой строки
    }

    // Обработка таблицы в конце текста
    if (tableLines.Count >= 2 && tableStart != -1)
    {
        string tableText = string.Join("\n", tableLines);
        var element = new MarkdownElementMatch
        {
            ElementType = MarkdownElementType.Table,
            StartPosition = tableStart,
            EndPosition = tableStart + tableText.Length,
            Content = tableText,
            FullMatch = tableText,
            Metadata = new Dictionary<string, object>
            {
                { "RowCount", tableLines.Count }
            }
        };

        matches.Add(element);
    }

    return matches;
}
```

**Проверка:** Убедитесь, что класс `MarkdownPatternMatcher` компилируется без ошибок.

---

## Этап 3: Реализация MarkdownElementFormatter

### Шаг 3.1: Создание класса MarkdownElementFormatter

**Действие:** Добавьте класс `MarkdownElementFormatter` в файл `WordMarkdownFormatter.cs`

```csharp
/// <summary>
/// Класс для применения форматирования Word к найденным элементам Markdown
/// </summary>
private class MarkdownElementFormatter
{
    private readonly Document _activeDoc;
    private readonly Application _wordApp;

    /// <summary>
    /// Конструктор
    /// </summary>
    public MarkdownElementFormatter(Document activeDoc, Application wordApp)
    {
        _activeDoc = activeDoc ?? throw new ArgumentNullException(nameof(activeDoc));
        _wordApp = wordApp ?? throw new ArgumentNullException(nameof(wordApp));
    }
}
```

### Шаг 3.2: Реализация метода FormatHeading

**Действие:** Добавьте метод для форматирования заголовков

```csharp
/// <summary>
/// Применение стиля заголовка к найденному элементу
/// </summary>
public void FormatHeading(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null || !element.Metadata.ContainsKey("Level"))
            return;

        int level = (int)element.Metadata["Level"];
        if (level < 1 || level > 6)
            return;

        // Получаем диапазон заголовка в документе
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range headingRange = _activeDoc.Range(start, end);

        // Определяем стиль заголовка
        WdBuiltinStyle headingStyle;
        switch (level)
        {
            case 1: headingStyle = WdBuiltinStyle.wdStyleHeading1; break;
            case 2: headingStyle = WdBuiltinStyle.wdStyleHeading2; break;
            case 3: headingStyle = WdBuiltinStyle.wdStyleHeading3; break;
            case 4: headingStyle = WdBuiltinStyle.wdStyleHeading4; break;
            case 5: headingStyle = WdBuiltinStyle.wdStyleHeading5; break;
            case 6: headingStyle = WdBuiltinStyle.wdStyleHeading6; break;
            default: headingStyle = WdBuiltinStyle.wdStyleNormal; break;
        }

        // Применяем стиль
        headingRange.set_Style(headingStyle);

        // Удаляем символы # и пробелы из начала
        string currentText = headingRange.Text;
        string newText = element.Content;
        
        // Заменяем текст, сохраняя форматирование
        headingRange.Text = newText;
        
        // Повторно применяем стиль после замены текста
        headingRange.set_Style(headingStyle);
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatHeading] Ошибка: {ex.Message}");
    }
}
```

### Шаг 3.3: Реализация методов форматирования текста

**Действие:** Добавьте методы для форматирования жирного, курсивного и зачеркнутого текста

```csharp
/// <summary>
/// Применение жирного форматирования
/// </summary>
public void FormatBoldText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем жирный шрифт
        textRange.Font.Bold = -1;

        // Удаляем символы **
        string currentText = textRange.Text;
        string newText = element.Content;
        textRange.Text = newText;
        
        // Повторно применяем форматирование
        textRange.Font.Bold = -1;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatBoldText] Ошибка: {ex.Message}");
    }
}

/// <summary>
/// Применение курсива
/// </summary>
public void FormatItalicText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем курсив
        textRange.Font.Italic = -1;

        // Удаляем символы *
        string newText = element.Content;
        textRange.Text = newText;
        
        // Повторно применяем форматирование
        textRange.Font.Italic = -1;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatItalicText] Ошибка: {ex.Message}");
    }
}

/// <summary>
/// Применение зачеркивания
/// </summary>
public void FormatStrikethroughText(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range textRange = _activeDoc.Range(start, end);

        // Устанавливаем зачеркивание
        textRange.Font.StrikeThrough = -1;

        // Удаляем символы ~~
        string newText = element.Content;
        textRange.Text = newText;
        
        // Повторно применяем форматирование
        textRange.Font.StrikeThrough = -1;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatStrikethroughText] Ошибка: {ex.Message}");
    }
}
```

### Шаг 3.4: Реализация методов форматирования кода

**Действие:** Добавьте методы для форматирования инлайн-кода и блоков кода

```csharp
/// <summary>
/// Применение форматирования для инлайн-кода
/// </summary>
public void FormatInlineCode(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range codeRange = _activeDoc.Range(start, end);

        // Устанавливаем моноширинный шрифт
        codeRange.Font.Name = "Courier New";
        codeRange.Font.Size = 10;

        // Опционально: добавляем фон
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;

        // Удаляем обратные кавычки
        string newText = element.Content;
        codeRange.Text = newText;
        
        // Повторно применяем форматирование
        codeRange.Font.Name = "Courier New";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatInlineCode] Ошибка: {ex.Message}");
    }
}

/// <summary>
/// Применение форматирования для блока кода
/// </summary>
public void FormatCodeBlock(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range codeRange = _activeDoc.Range(start, end);

        // Устанавливаем моноширинный шрифт
        codeRange.Font.Name = "Consolas";
        codeRange.Font.Size = 10;

        // Добавляем фон и отступы
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        codeRange.ParagraphFormat.LeftIndent = 18;
        codeRange.ParagraphFormat.RightIndent = 18;
        codeRange.ParagraphFormat.SpaceBefore = 6;
        codeRange.ParagraphFormat.SpaceAfter = 6;

        // Удаляем символы ```
        string newText = element.Content;
        codeRange.Text = newText;
        
        // Повторно применяем форматирование
        codeRange.Font.Name = "Consolas";
        codeRange.Font.Size = 10;
        codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        codeRange.ParagraphFormat.LeftIndent = 18;
        codeRange.ParagraphFormat.RightIndent = 18;
        codeRange.ParagraphFormat.SpaceBefore = 6;
        codeRange.ParagraphFormat.SpaceAfter = 6;
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatCodeBlock] Ошибка: {ex.Message}");
    }
}
```

### Шаг 3.5: Реализация методов форматирования ссылок и списков

**Действие:** Добавьте методы для форматирования ссылок и элементов списков

```csharp
/// <summary>
/// Создание гиперссылки из Markdown-ссылки
/// </summary>
public void FormatLink(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null || !element.Metadata.ContainsKey("Url"))
            return;

        string url = element.Metadata["Url"].ToString();
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range linkRange = _activeDoc.Range(start, end);

        // Заменяем текст [текст](url) на текст
        string linkText = element.Content;
        linkRange.Text = linkText;

        // Создаем гиперссылку
        try
        {
            Hyperlink hyperlink = _activeDoc.Hyperlinks.Add(
                linkRange,
                url,
                null, // Anchor
                null, // ScreenTip
                linkText // TextToDisplay
            );

            // Применяем подчеркивание
            linkRange.Font.Underline = WdUnderline.wdUnderlineSingle;
        }
        catch
        {
            // Если не удалось создать гиперссылку, просто подчеркиваем текст
            linkRange.Font.Underline = WdUnderline.wdUnderlineSingle;
            linkRange.Text = $"{linkText} ({url})";
        }
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatLink] Ошибка: {ex.Message}");
    }
}

/// <summary>
/// Применение форматирования списка
/// </summary>
public void FormatListItem(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range listRange = _activeDoc.Range(start, end);

        // Получаем параграф, содержащий элемент списка
        Paragraph paragraph = listRange.Paragraphs[1];
        
        // Определяем тип списка
        bool isOrdered = element.Metadata.ContainsKey("IsOrdered") && 
                        (bool)element.Metadata["IsOrdered"];

        // Удаляем маркер списка из текста
        string newText = element.Content;
        paragraph.Range.Text = newText;

        // Применяем форматирование списка
        if (isOrdered)
        {
            paragraph.Range.ListFormat.ApplyNumberDefault();
        }
        else
        {
            paragraph.Range.ListFormat.ApplyBulletDefault();
        }
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatListItem] Ошибка: {ex.Message}");
    }
}
```

### Шаг 3.6: Реализация методов форматирования цитат и таблиц

**Действие:** Добавьте оставшиеся методы форматирования

```csharp
/// <summary>
/// Применение стиля цитаты
/// </summary>
public void FormatQuote(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range quoteRange = _activeDoc.Range(start, end);

        // Получаем параграф
        Paragraph paragraph = quoteRange.Paragraphs[1];

        // Удаляем символ >
        string newText = element.Content;
        paragraph.Range.Text = newText;

        // Применяем стиль цитаты
        try
        {
            paragraph.Range.set_Style("Quote");
        }
        catch
        {
            // Если стиль не существует, используем Normal с отступом
            paragraph.Range.set_Style(WdBuiltinStyle.wdStyleNormal);
            paragraph.Range.ParagraphFormat.LeftIndent = 36; // 0.5 дюйма
        }
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatQuote] Ошибка: {ex.Message}");
    }
}

/// <summary>
/// Создание таблицы Word из Markdown-таблицы
/// ВАЖНО: Это упрощенная реализация. Полная реализация требует парсинга структуры таблицы.
/// </summary>
public void FormatTable(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null)
            return;

        // Упрощенная реализация - удаляем Markdown-синтаксис таблицы
        // Полная реализация должна парсить строки и столбцы и создавать таблицу Word
        int start = documentRange.Start + element.StartPosition;
        int end = documentRange.Start + element.EndPosition;
        Range tableRange = _activeDoc.Range(start, end);

        // Удаляем Markdown-синтаксис таблицы
        // В полной реализации здесь должен быть код создания таблицы Word
        Debug.WriteLine($"[FormatTable] Упрощенная реализация - таблица не преобразована");
        
        // Для полной реализации нужно:
        // 1. Парсить строки таблицы (разделитель |)
        // 2. Определить количество столбцов
        // 3. Создать таблицу Word через Tables.Add()
        // 4. Заполнить ячейки данными
        // 5. Применить форматирование
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"[FormatTable] Ошибка: {ex.Message}");
    }
}
```

**Проверка:** Убедитесь, что класс `MarkdownElementFormatter` компилируется без ошибок.

---

## Этап 4: Реализация WordMarkdownFormatter

### Шаг 4.1: Добавление приватных полей

**Действие:** Добавьте приватные поля в класс `WordMarkdownFormatter`

```csharp
public class WordMarkdownFormatter
{
    // Константы
    private const int MaxDocumentSize = 10000000;  // 10 миллионов символов
    private const int MaxPatternMatches = 10000;   // Максимум найденных элементов

    // Приватные поля
    private readonly Application _wordApp;
    private readonly Document _activeDoc;
    private readonly MarkdownPatternMatcher _patternMatcher;
    private readonly MarkdownElementFormatter _elementFormatter;
    private readonly List<MarkdownElementType> _processingOrder;
    private bool _isProcessing = false;

    /// <summary>
    /// Конструктор класса WordMarkdownFormatter
    /// </summary>
    public WordMarkdownFormatter()
    {
        _wordApp = Globals.ThisAddIn.Application;
        _activeDoc = _wordApp?.ActiveDocument;

        if (_activeDoc == null)
            throw new InvalidOperationException("Нет активного документа Word.");

        _patternMatcher = new MarkdownPatternMatcher();
        _elementFormatter = new MarkdownElementFormatter(_activeDoc, _wordApp);

        // Определяем порядок обработки элементов (от более специфичных к менее специфичным)
        _processingOrder = new List<MarkdownElementType>
        {
            MarkdownElementType.CodeBlock,      // Сначала блоки кода (многострочные)
            MarkdownElementType.Table,           // Затем таблицы
            MarkdownElementType.Heading,        // Затем заголовки
            MarkdownElementType.Quote,           // Затем цитаты
            MarkdownElementType.ListItem,        // Затем списки
            MarkdownElementType.Link,            // Затем ссылки
            MarkdownElementType.Bold,            // Затем форматирование текста
            MarkdownElementType.Italic,
            MarkdownElementType.Strikethrough,
            MarkdownElementType.InlineCode,
            MarkdownElementType.HorizontalRule  // В конце горизонтальные линии
        };
    }
}
```

### Шаг 4.2: Реализация метода ValidateDocument

**Действие:** Добавьте метод валидации документа

```csharp
/// <summary>
/// Валидация состояния документа перед форматированием
/// </summary>
private bool ValidateDocument()
{
    try
    {
        if (_activeDoc == null)
        {
            LogError("Активный документ не найден.");
            return false;
        }

        // Проверка на защиту от изменений
        if (_activeDoc.ProtectionType != WdProtectionType.wdNoProtection)
        {
            LogError("Документ защищен от изменений.");
            return false;
        }

        // Проверка размера документа
        int documentSize = _activeDoc.Content.Text.Length;
        if (documentSize > MaxDocumentSize)
        {
            LogError($"Размер документа ({documentSize}) превышает максимально допустимый ({MaxDocumentSize}).");
            return false;
        }

        return true;
    }
    catch (Exception ex)
    {
        LogError("Ошибка при валидации документа.", ex);
        return false;
    }
}
```

### Шаг 4.3: Реализация метода FindAllMarkdownElements

**Действие:** Добавьте метод поиска всех элементов Markdown

```csharp
/// <summary>
/// Поиск всех элементов Markdown в указанном диапазоне
/// </summary>
private List<MarkdownElementMatch> FindAllMarkdownElements(Range range)
{
    var allElements = new List<MarkdownElementMatch>();

    try
    {
        if (range == null)
            return allElements;

        // Извлекаем текст из диапазона
        string text = range.Text;

        if (string.IsNullOrEmpty(text))
            return allElements;

        // Поиск всех типов элементов
        allElements.AddRange(_patternMatcher.FindCodeBlocks(text));
        allElements.AddRange(_patternMatcher.FindTables(text));
        allElements.AddRange(_patternMatcher.FindHeadings(text));
        allElements.AddRange(_patternMatcher.FindQuotes(text));
        allElements.AddRange(_patternMatcher.FindListItems(text));
        allElements.AddRange(_patternMatcher.FindLinks(text));
        allElements.AddRange(_patternMatcher.FindBoldText(text));
        allElements.AddRange(_patternMatcher.FindItalicText(text));
        allElements.AddRange(_patternMatcher.FindStrikethroughText(text));
        allElements.AddRange(_patternMatcher.FindInlineCode(text));
        allElements.AddRange(_patternMatcher.FindHorizontalRules(text));

        // Проверка на превышение лимита
        if (allElements.Count > MaxPatternMatches)
        {
            LogError($"Найдено слишком много элементов ({allElements.Count}). Обработка может быть медленной.");
            // Продолжаем обработку, но предупреждаем
        }

        // Сортировка по позиции (от начала к концу)
        allElements = allElements.OrderBy(e => e.StartPosition).ToList();

        // Фильтрация перекрывающихся элементов
        // Приоритет отдается более длинным или более специфичным элементам
        allElements = FilterOverlappingElements(allElements);

        return allElements;
    }
    catch (Exception ex)
    {
        LogError("Ошибка при поиске элементов Markdown.", ex);
        return allElements;
    }
}

/// <summary>
/// Фильтрация перекрывающихся элементов
/// </summary>
private List<MarkdownElementMatch> FilterOverlappingElements(List<MarkdownElementMatch> elements)
{
    if (elements == null || elements.Count == 0)
        return elements;

    var filtered = new List<MarkdownElementMatch>();
    var processed = new HashSet<int>();

    // Сортируем по приоритету (согласно _processingOrder) и длине
    var sorted = elements.OrderByDescending(e => 
    {
        int priority = _processingOrder.IndexOf(e.ElementType);
        return priority >= 0 ? priority : int.MaxValue;
    }).ThenByDescending(e => e.Length).ToList();

    foreach (var element in sorted)
    {
        // Проверяем, не перекрывается ли элемент с уже обработанными
        bool overlaps = false;
        for (int i = element.StartPosition; i < element.EndPosition; i++)
        {
            if (processed.Contains(i))
            {
                overlaps = true;
                break;
            }
        }

        if (!overlaps)
        {
            filtered.Add(element);
            for (int i = element.StartPosition; i < element.EndPosition; i++)
            {
                processed.Add(i);
            }
        }
    }

    // Возвращаем отсортированные по позиции
    return filtered.OrderBy(e => e.StartPosition).ToList();
}
```

### Шаг 4.4: Реализация метода ApplyFormattingToElement

**Действие:** Добавьте метод применения форматирования к элементу

```csharp
/// <summary>
/// Применение форматирования к одному найденному элементу
/// </summary>
private void ApplyFormattingToElement(MarkdownElementMatch element, Range documentRange)
{
    try
    {
        if (element == null || documentRange == null)
            return;

        switch (element.ElementType)
        {
            case MarkdownElementType.Heading:
                _elementFormatter.FormatHeading(element, documentRange);
                break;

            case MarkdownElementType.Bold:
                _elementFormatter.FormatBoldText(element, documentRange);
                break;

            case MarkdownElementType.Italic:
                _elementFormatter.FormatItalicText(element, documentRange);
                break;

            case MarkdownElementType.Strikethrough:
                _elementFormatter.FormatStrikethroughText(element, documentRange);
                break;

            case MarkdownElementType.InlineCode:
                _elementFormatter.FormatInlineCode(element, documentRange);
                break;

            case MarkdownElementType.CodeBlock:
                _elementFormatter.FormatCodeBlock(element, documentRange);
                break;

            case MarkdownElementType.Link:
                _elementFormatter.FormatLink(element, documentRange);
                break;

            case MarkdownElementType.ListItem:
                _elementFormatter.FormatListItem(element, documentRange);
                break;

            case MarkdownElementType.Quote:
                _elementFormatter.FormatQuote(element, documentRange);
                break;

            case MarkdownElementType.Table:
                _elementFormatter.FormatTable(element, documentRange);
                break;

            case MarkdownElementType.HorizontalRule:
                // Горизонтальные линии можно обработать отдельно
                // Например, вставить разделитель или пустой параграф
                break;

            default:
                Debug.WriteLine($"[ApplyFormattingToElement] Неизвестный тип элемента: {element.ElementType}");
                break;
        }
    }
    catch (Exception ex)
    {
        LogError($"Ошибка при применении форматирования к элементу {element?.ElementType}.", ex);
    }
}
```

### Шаг 4.5: Реализация основных публичных методов

**Действие:** Добавьте публичные методы `FormatMarkdownInWord`, `FormatSelectedText`, `FormatEntireDocument` и `CanFormat`

```csharp
/// <summary>
/// Главный метод модуля, запускающий процесс форматирования
/// </summary>
/// <param name="targetRange">Диапазон текста для обработки. Если null, обрабатывается весь документ</param>
public void FormatMarkdownInWord(Range targetRange = null)
{
    if (_isProcessing)
    {
        LogError("Процесс форматирования уже запущен.");
        return;
    }

    try
    {
        _isProcessing = true;

        // Валидация документа
        if (!ValidateDocument())
        {
            return;
        }

        // Определение диапазона обработки
        Range rangeToProcess = targetRange ?? _activeDoc.Content;

        if (rangeToProcess == null)
        {
            LogError("Не удалось определить диапазон для обработки.");
            return;
        }

        // Проверка размера диапазона
        int rangeSize = rangeToProcess.Text.Length;
        if (rangeSize > MaxDocumentSize)
        {
            LogError($"Размер диапазона ({rangeSize}) превышает максимально допустимый ({MaxDocumentSize}).");
            return;
        }

        // Поиск всех элементов Markdown
        var elements = FindAllMarkdownElements(rangeToProcess);

        if (elements == null || elements.Count == 0)
        {
            Debug.WriteLine("[FormatMarkdownInWord] Элементы Markdown не найдены.");
            return;
        }

        // Сортировка элементов по приоритету обработки
        var sortedElements = elements.OrderBy(e =>
        {
            int priority = _processingOrder.IndexOf(e.ElementType);
            return priority >= 0 ? priority : int.MaxValue;
        }).ToList();

        // Применение форматирования к каждому элементу
        // ВАЖНО: Обрабатываем в обратном порядке (с конца к началу),
        // чтобы позиции не сдвигались при удалении синтаксиса
        foreach (var element in sortedElements.OrderByDescending(e => e.StartPosition))
        {
            ApplyFormattingToElement(element, rangeToProcess);
        }

        Debug.WriteLine($"[FormatMarkdownInWord] Обработано {sortedElements.Count} элементов Markdown.");
    }
    catch (Exception ex)
    {
        LogError("Критическая ошибка при форматировании Markdown.", ex);
    }
    finally
    {
        ResetProcessingState();
    }
}

/// <summary>
/// Удобный метод для форматирования только выделенного пользователем текста
/// </summary>
public void FormatSelectedText()
{
    try
    {
        Range selection = _wordApp.Selection?.Range;
        if (selection == null)
        {
            LogError("Не удалось получить выделенный текст.");
            return;
        }

        if (selection.Text.Length == 0)
        {
            LogError("Текст не выделен.");
            return;
        }

        FormatMarkdownInWord(selection);
    }
    catch (Exception ex)
    {
        LogError("Ошибка при форматировании выделенного текста.", ex);
    }
}

/// <summary>
/// Удобный метод для форматирования всего документа
/// </summary>
public void FormatEntireDocument()
{
    try
    {
        Range entireDocument = _activeDoc.Content;
        FormatMarkdownInWord(entireDocument);
    }
    catch (Exception ex)
    {
        LogError("Ошибка при форматировании всего документа.", ex);
    }
}

/// <summary>
/// Проверка возможности выполнения форматирования
/// </summary>
public bool CanFormat()
{
    try
    {
        if (_activeDoc == null)
            return false;

        if (_activeDoc.ProtectionType != WdProtectionType.wdNoProtection)
            return false;

        if (_isProcessing)
            return false;

        int documentSize = _activeDoc.Content.Text.Length;
        if (documentSize > MaxDocumentSize)
            return false;

        return true;
    }
    catch
    {
        return false;
    }
}
```

### Шаг 4.6: Реализация вспомогательных методов

**Действие:** Добавьте методы логирования и сброса состояния

```csharp
/// <summary>
/// Логирование ошибок в процессе форматирования
/// </summary>
private void LogError(string message, Exception exception = null)
{
    string logMessage = $"[WordMarkdownFormatter] {message}";
    
    if (exception != null)
    {
        logMessage += $"\nИсключение: {exception.Message}\nСтек вызовов: {exception.StackTrace}";
    }

    Debug.WriteLine(logMessage);
    
    // Здесь можно добавить запись в файл лога, если требуется
}

/// <summary>
/// Сброс состояния обработки после завершения или ошибки
/// </summary>
private void ResetProcessingState()
{
    _isProcessing = false;
    // Здесь можно добавить очистку других временных данных
}
```

**Проверка:** Убедитесь, что класс `WordMarkdownFormatter` компилируется без ошибок.

---

## Этап 5: Интеграция с Ribbon

### Шаг 5.1: Добавление кнопки в Ribbon Designer

**Действие:** Откройте файл `MarkdownRibbon.Designer.cs` и добавьте новую кнопку в группу `grpFormat`

**Найдите секцию с группой `grpFormat` и добавьте:**

```csharp
// В секции объявления полей (около строки 291):
internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatMarkdown;

// В методе InitializeComponent, в секции grpFormat (после bCode):
// 
// btnFormatMarkdown
// 
this.btnFormatMarkdown = this.Factory.CreateRibbonButton();
this.btnFormatMarkdown.Label = "Форматировать Markdown";
this.btnFormatMarkdown.Name = "btnFormatMarkdown";
this.btnFormatMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatMarkdown_Click);

// В секции grpFormat.Items.Add (после this.grpFormat.Items.Add(this.bCode);):
this.grpFormat.Items.Add(this.btnFormatMarkdown);
```

### Шаг 5.2: Добавление обработчика события

**Действие:** Откройте файл `MarkdownRibbon.cs` и добавьте обработчик события

```csharp
private void btnFormatMarkdown_Click(object sender, RibbonControlEventArgs e)
{
    try
    {
        var formatter = new Services.WordMarkdownFormatter();
        
        // Проверяем, есть ли выделенный текст
        Range selection = Globals.ThisAddIn.Application.Selection?.Range;
        if (selection != null && selection.Text.Length > 0)
        {
            // Форматируем только выделенный текст
            formatter.FormatSelectedText();
            MessageBox.Show(
                "Markdown-синтаксис в выделенном тексте успешно отформатирован!",
                "Успех",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        else
        {
            // Форматируем весь документ
            var result = MessageBox.Show(
                "Текст не выделен. Отформатировать весь документ?",
                "Подтверждение",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                formatter.FormatEntireDocument();
                MessageBox.Show(
                    "Markdown-синтаксис в документе успешно отформатирован!",
                    "Успех",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show(
            $"Ошибка при форматировании Markdown: {ex.Message}",
            "Ошибка",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
}
```

**Проверка:** Убедитесь, что проект компилируется и кнопка появляется в Ribbon.

---

## Этап 6: Тестирование

### Шаг 6.1: Создание тестового документа

**Действие:** Создайте тестовый документ Word с различными элементами Markdown:

```
# Заголовок 1

## Заголовок 2

Это **жирный текст** и это *курсивный текст*.

Также есть ~~зачеркнутый текст~~.

Вот `инлайн-код` в тексте.

- Элемент списка 1
- Элемент списка 2
- Элемент списка 3

1. Нумерованный элемент 1
2. Нумерованный элемент 2

> Это цитата
> из нескольких строк

[Ссылка на Google](https://www.google.com)

---

```python
def hello():
    print("Hello, World!")
```

| Заголовок 1 | Заголовок 2 |
|-------------|-------------|
| Ячейка 1    | Ячейка 2    |
```

### Шаг 6.2: Тестирование базовой функциональности

**Действия:**
1. Откройте тестовый документ в Word
2. Нажмите кнопку "Форматировать Markdown" в Ribbon
3. Проверьте, что:
   - Заголовки получили соответствующие стили
   - Жирный текст стал жирным
   - Курсивный текст стал курсивом
   - Зачеркнутый текст стал зачеркнутым
   - Инлайн-код получил моноширинный шрифт
   - Списки стали списками Word
   - Цитаты получили стиль цитаты
   - Ссылки стали гиперссылками
   - Синтаксис Markdown удален

### Шаг 6.3: Тестирование выделенного текста

**Действия:**
1. Выделите часть текста с Markdown-синтаксисом
2. Нажмите кнопку "Форматировать Markdown"
3. Проверьте, что форматирование применилось только к выделенному тексту

### Шаг 6.4: Тестирование обработки ошибок

**Действия:**
1. Попробуйте отформатировать защищенный документ
2. Попробуйте отформатировать очень большой документ
3. Проверьте, что ошибки обрабатываются корректно

---

## Приложение: Полный код классов

Полный код всех классов будет слишком большим для одного файла. Рекомендуется разбить на несколько файлов:

1. **Services/WordMarkdownFormatter.cs** - основной класс
2. **Services/MarkdownPatternMatcher.cs** - класс поиска (опционально, можно оставить вложенным)
3. **Services/MarkdownElementFormatter.cs** - класс форматирования (опционально, можно оставить вложенным)

Или все классы можно разместить в одном файле `Services/WordMarkdownFormatter.cs` с вложенными классами.

---

## Заключение

После выполнения всех этапов модуль `WordMarkdownFormatter` будет готов к использованию. Модуль позволит автоматически форматировать Markdown-синтаксис в документах Word, применяя соответствующие стили и удаляя синтаксические маркеры.

### Следующие шаги (опционально):
1. Добавление поддержки дополнительных элементов Markdown
2. Улучшение парсинга таблиц
3. Добавление настроек пользователя
4. Оптимизация производительности для больших документов
5. Добавление асинхронной обработки

---

## Примечания по реализации

### Важные замечания:

1. **Обработка позиций**: При удалении синтаксиса позиции элементов сдвигаются. Поэтому важно обрабатывать элементы в обратном порядке (с конца к началу).

2. **Перекрывающиеся элементы**: Некоторые элементы Markdown могут перекрываться (например, жирный текст внутри заголовка). Алгоритм фильтрации должен правильно обрабатывать такие случаи.

3. **COM-объекты**: При работе с Word Interop важно правильно обрабатывать COM-объекты и освобождать ресурсы.

4. **Производительность**: Для больших документов может потребоваться оптимизация или асинхронная обработка.

5. **Таблицы**: Полная реализация парсинга таблиц требует более сложной логики, чем представлено в упрощенной версии.

---

**Дата создания инструкции:** 2024
**Версия:** 1.0




