using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordMarkdownAddIn.Services
{
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

    /// <summary>
    /// Класс для поиска и распознавания синтаксиса Markdown в тексте документа
    /// </summary>
    internal class MarkdownPatternMatcher
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
    }

    /// <summary>
    /// Класс для применения форматирования Word к найденным элементам Markdown
    /// </summary>
    internal class MarkdownElementFormatter
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

        /// <summary>
        /// Логирование ошибок в процессе форматирования
        /// </summary>
        private void LogError(string message, Exception exception = null)
        {
            string logMessage = $"[MarkdownElementFormatter] {message}";

            if (exception != null)
            {
                logMessage += $"\nИсключение: {exception.Message}\nСтек вызовов: {exception.StackTrace}";
            }

            Debug.WriteLine(logMessage);
        }
    }



    /// <summary>
    /// Вспомогательный класс для хранения контекста форматирования во время обработки
    /// </summary>
    internal class FormattingContext
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

    /// <summary>
    /// Основной класс для форматирования Markdown-синтаксиса в документе Word.
    /// Находит элементы Markdown, применяет форматирование Word и удаляет синтаксические маркеры.
    /// </summary>
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
    }

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



  
}
