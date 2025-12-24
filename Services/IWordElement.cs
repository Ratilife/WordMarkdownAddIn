using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using WordMarkdownAddIn.Services; // Для доступа к SyntaxHighlighter

namespace WordMarkdownAddIn.Services
{
    public interface IWordElement
    {
        string ElementType { get; }
        string ToMarkdown(); // Метод для преобразования элемента в строку Markdown
        void ApplyToWord(Document doc); // Word ← Markdown
    }

    public interface IListItemElement : IWordElement { }

    // Класс для представления фрагмента текста с единым форматированием
    public class FormattedRun
    {
        public string ElementType => "Text";
        public string Text { get; set; } = "";      //Сам текст
        public bool IsBold { get; set; }            // Жирный шрифт
        public bool IsItalic { get; set; }          // Курсив
        public bool IsStrikethrough { get; set; }   // Зачеркнутый
        public bool IsUnderline { get; set; }       // Подчеркивание
        public bool IsSuperscript { get; set; }     //Надстрочный текст (например, x²)
        public bool IsSubscript { get; set; }       // Подстрочный текст (например, H₂O)
        public bool SmallCaps { get; set; }         // Уменьшенные заглавные буквы
        public bool AllCaps { get; set; }           // ВСЁ ЗАГЛАВНЫМИ (все буквы заглавные)
    }


    public class WordFormattedText : IWordElement
    {
        public string ElementType => "Content";
        public List<FormattedRun> Runs { get; set; } = new List<FormattedRun>();
        // можно добавить формтирование всего блока (блочное форматирование)
        public string ToMarkdown()
        {
            if (Runs == null || Runs.Count == 0)
                return "";

            var sb = new StringBuilder();
            foreach (var run in Runs)
            {
                if (run == null || string.IsNullOrEmpty(run.Text))
                    continue;

                string text = run.Text;
                string markdown = text;

                // Применяем форматирование в правильном порядке
                // Сначала применяем подстрочный/надстрочный текст к исходному тексту
                if (run.IsSubscript)
                    markdown = $"<sub>{text}</sub>";
                else if (run.IsSuperscript)
                    markdown = $"<sup>{text}</sup>";

                // Затем применяем Markdown форматирование
                if (run.IsStrikethrough)
                    markdown = $"~~{markdown}~~";

                if (run.IsBold)
                    markdown = $"**{markdown}**";

                if (run.IsItalic)
                    markdown = $"*{markdown}*";

                // Подчеркивание (в Markdown нет нативного синтаксиса, используем HTML)
                // Применяем последним, чтобы обернуть все форматирование
                if (run.IsUnderline)
                    markdown = $"<u>{markdown}</u>";

                sb.Append(markdown);
            }

            return sb.ToString();
        }
        public void ApplyToWord(Document doc, Range range)
        {
            // НЕ создаем новый параграф, работаем с переданным Range
            foreach (var run in Runs)
            {
                if (run == null || string.IsNullOrEmpty(run.Text))
                    continue;

                // В Word Interop надёжнее сначала вставить текст, затем применить форматирование
                // к конкретному диапазону вставленного фрагмента.

                // 1) Запоминаем позицию начала вставки
                int start = range.End;

                // Проверяем, находится ли Range внутри таблицы
                bool isInTable = false;
                try
                {
                    isInTable = range.Information[WdInformation.wdWithInTable];
                }
                catch
                {
                    // Если проверка не удалась, считаем что не в таблице
                    isInTable = false;
                }

                // Очищаем форматирование списка перед вставкой текста (только если НЕ в таблице)
                if (!isInTable)
                {
                    try
                    {
                        range.ListFormat.RemoveNumbers();
                    }
                    catch { } // Игнорируем ошибки
                }

                // 2) Вставляем текст в конец текущего range
                int end;
                if (isInTable)
                {
                    // Для ячеек таблицы используем безопасный метод вставки
                    try
                    {
                        // Получаем текущий текст без маркера конца ячейки
                        int rangeStart = range.Start;
                        int rangeEnd = range.End;
                        
                        // Создаем Range без маркера конца ячейки
                        if (rangeEnd > rangeStart)
                        {
                            rangeEnd--; // Исключаем маркер
                        }
                        
                        if (rangeEnd > rangeStart)
                        {
                            Range safeRange = doc.Range(rangeStart, rangeEnd);
                            string currentText = safeRange.Text ?? "";
                            safeRange.Text = currentText + run.Text;
                            end = safeRange.End;
                            // Обновляем исходный range
                            range.SetRange(safeRange.Start, safeRange.End);
                        }
                        else
                        {
                            // Если ячейка пустая, создаем минимальный Range и вставляем текст
                            try
                            {
                                // Создаем Range для пустой ячейки (без маркера)
                                Range emptyRange = doc.Range(rangeStart, rangeStart);
                                emptyRange.Text = run.Text;
                                end = emptyRange.End;
                                range.SetRange(emptyRange.Start, emptyRange.End);
                            }
                            catch
                            {
                                // Если не удалось, просто пропускаем
                                end = range.End;
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        // Если все методы не работают, просто пропускаем этот run
                        end = range.End;
                        continue; // Пропускаем этот run
                    }
                }
                else
                {
                    try
                    {
                        range.InsertAfter(run.Text);
                        end = start + run.Text.Length;
                    }
                    catch
                    {
                        // Если InsertAfter не работает, используем прямое присваивание
                        string currentText = range.Text ?? "";
                        range.Text = currentText + run.Text;
                        end = range.End;
                    }
                }

                // 3) Получаем диапазон вставленного текста [start, end)
                var insertedRange = doc.Range(start, end);

                // 4) Применяем форматирование к вставленному диапазону
                insertedRange.Font.Bold = run.IsBold ? -1 : 0;
                insertedRange.Font.Italic = run.IsItalic ? -1 : 0;
                insertedRange.Font.StrikeThrough = run.IsStrikethrough ? -1 : 0;
                insertedRange.Font.Underline = run.IsUnderline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;

                insertedRange.Font.Superscript = run.IsSuperscript ? -1 : 0;
                insertedRange.Font.Subscript = run.IsSubscript ? -1 : 0;
                insertedRange.Font.SmallCaps = run.SmallCaps ? -1 : 0;
                insertedRange.Font.AllCaps = run.AllCaps ? -1 : 0;

                // 5) Сдвигаем исходный range в конец, чтобы следующий run вставлялся после этого текста
                range.SetRange(end, end);
            }
        }

        // Реализация интерфейса IWordElement - создает новый параграф
        public void ApplyToWord(Document doc)
        {
            if (doc == null)
                return;

            // Создаем новый параграф
            var paragraph = doc.Content.Paragraphs.Add();

            // Применяем форматированный текст к параграфу через перегрузку с Range
            ApplyToWord(doc, paragraph.Range);

            // Добавляем перенос строки
            paragraph.Range.InsertParagraphAfter();
        }
    }

    public class WordTable : IWordElement
    {
        public string ElementType => "Table";
        public List<List<WordFormattedText>> Rows { get; set; } = new List<List<WordFormattedText>>();

        public WordTable(List<List<WordFormattedText>> rows)
        {
            Rows = rows ?? new List<List<WordFormattedText>>(); // защита от null
        }

        // Вспомогательный метод для нормализации содержимого ячейки
        private string NormalizeCellContent(string cellContent)
        {
            if (string.IsNullOrEmpty(cellContent))
                return "";

            // Заменяем переносы строк на пробелы
            cellContent = cellContent.Replace("\r\n", " ")
                                     .Replace("\n", " ")
                                     .Replace("\r", " ");

            // Убираем множественные пробелы
            cellContent = System.Text.RegularExpressions.Regex.Replace(cellContent, @"\s+", " ");

            // Экранируем символы |
            cellContent = cellContent.Replace("|", "\\|");

            return cellContent.Trim();
        }

        public string ToMarkdown()
        {
            if (Rows == null || Rows.Count == 0)
                return "";

            // Находим максимальное количество колонок
            int maxColumns = 0;
            foreach (var row in Rows)
            {
                if (row != null && row.Count > maxColumns)
                {
                    maxColumns = row.Count;
                }
            }

            // Если нет колонок, возвращаем пустую строку
            if (maxColumns == 0)
                return "";

            var sb = new StringBuilder();

            // Обрабатываем первую строку как заголовок таблицы
            if (Rows.Count > 0)
            {
                var headerRow = Rows[0];
                if (headerRow != null && headerRow.Count > 0)
                {
                    //  Нормализуем заголовок - добавляем пустые ячейки если нужно
                    while (headerRow.Count < maxColumns)
                    {
                        headerRow.Add(new WordFormattedText()); // Добавляем пустую ячейку
                    }

                    // Заголовок таблицы
                    sb.Append("| ");
                    for (int i = 0; i < maxColumns; i++)
                    {
                        var cell = i < headerRow.Count ? headerRow[i] : null;
                        string cellContent = cell?.ToMarkdown() ?? "";
                        // Обработка переносов строк (из Проблемы 2)
                        cellContent = NormalizeCellContent(cellContent);
                        sb.Append(cellContent.Replace("|", "\\|")); // Экранируем символы |
                        sb.Append(" | ");
                    }
                    sb.AppendLine();

                    // Разделитель заголовка
                    sb.Append("| ");
                    for (int i = 0; i < maxColumns; i++)
                    {
                        sb.Append(" --- |");
                    }
                    sb.AppendLine();
                }

                // Остальные строки
                for (int i = 1; i < Rows.Count; i++)
                {
                    var row = Rows[i];
                    if (row != null && row.Count > 0)
                    {
                        // Нормализуем каждую строку данных
                        while (row.Count < maxColumns)
                        {
                            row.Add(new WordFormattedText()); // Добавляем пустую ячейку
                        }

                        sb.Append("| ");
                        for (int j = 0; j < maxColumns; j++)
                        {
                            var cell = j < row.Count ? row[j] : null;
                            string cellContent = cell?.ToMarkdown() ?? "";
                            // Обработка переносов строк (из Проблемы 2)
                            cellContent = NormalizeCellContent(cellContent);
                            sb.Append(cellContent);
                            sb.Append(" | ");
                        }
                        sb.AppendLine();
                    }
                }
            }

            return sb.ToString().TrimEnd();
        }

        public void ApplyToWord(Document doc)
        {
            if (doc == null || Rows == null || Rows.Count == 0)
                return;

            // 1. Определяем размеры таблицы
            int rowCount = Rows.Count;
            int columnCount = Rows[0]?.Count ?? 0;

            if (columnCount == 0)
                return;

            // 2. Получаем позицию для вставки (в конец документа)
            var range = doc.Content;
            range.Collapse(WdCollapseDirection.wdCollapseEnd);

            // 3. Создаем таблицу
            var wordTable = doc.Tables.Add(
                range,
                rowCount,
                columnCount,
                WdDefaultTableBehavior.wdWord9TableBehavior,
                WdAutoFitBehavior.wdAutoFitFixed);

            // 4. Заполняем ячейки
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = Rows[rowIndex];
                if (row == null)
                    continue;

                for (int colIndex = 0; colIndex < columnCount && colIndex < row.Count; colIndex++)
                {
                    var cell = wordTable.Cell(rowIndex + 1, colIndex + 1); // Word индексация с 1
                    var cellContent = row[colIndex];

                    // Создаем Range без маркера конца ячейки
                    int cellStart = cell.Range.Start;
                    int cellEnd = cell.Range.End;

                    // Исключаем последний символ (маркер конца ячейки)
                    if (cellEnd > cellStart)
                    {
                        cellEnd--; // Уменьшаем на 1, чтобы исключить маркер
                    }

                    // Создаем чистый Range без маркера конца ячейки
                    Range cleanCellRange = doc.Range(cellStart, cellEnd);

                    // Проверяем, что Range создан успешно и валиден
                    if (cleanCellRange == null || cellEnd <= cellStart)
                    {
                        // Если Range невалиден, используем оригинальный cell.Range
                        cleanCellRange = cell.Range;
                        // Но все равно нужно исключить маркер - создаем Range вручную
                        if (cell.Range.End > cell.Range.Start)
                        {
                            int safeEnd = cell.Range.End - 1;
                            if (safeEnd > cell.Range.Start)
                            {
                                cleanCellRange = doc.Range(cell.Range.Start, safeEnd);
                            }
                            else
                            {
                                cleanCellRange = cell.Range; // Если ячейка пустая, используем как есть
                            }
                        }
                    }

                    if (cellContent != null)
                    {
                        // Применяем форматированный текст к ячейке
                        // Метод ApplyToWord уже правильно обрабатывает ячейки таблиц
                        cellContent.ApplyToWord(doc, cleanCellRange);
                    }
                    // Убираем избыточную очистку текста - метод ApplyToWord уже обрабатывает это правильно
                }
            }

            // 5. Форматируем таблицу (опционально)
            wordTable.AutoFormat(
                WdTableFormat.wdTableFormatGrid1,
                true, false, false, false, false, false, false);

            // 6. Добавляем перенос строки после таблицы
            range.InsertParagraphAfter();
        }
    }

    public class WordListItem : IWordElement
    {
        public string ElementType => "ListItem";
        public List<WordFormattedText> Contents { get; set; } // Неправильно для одного элемента списка
        public bool IsOrdered { get; set; }

        public WordListItem(List<WordFormattedText> contents, bool isOrdered)
        {
            Contents = contents;
            IsOrdered = isOrdered;
        }

        public string ToMarkdown()
        {
            if (Contents == null || Contents.Count == 0)
                return "";

            var sb = new StringBuilder();
            foreach (var content in Contents)
            {
                if (content != null)
                {
                    string markdown = content.ToMarkdown();
                    if (!string.IsNullOrEmpty(markdown))
                    {
                        sb.Append(markdown);
                    }
                }
            }

            return sb.ToString();
        }
        public void ApplyToWord(Document doc)
        {
            if (doc == null || Contents == null || Contents.Count == 0)
                return;

            // Обрабатываем каждый параграф внутри элемента списка
            foreach (var content in Contents)
            {
                if (content == null)
                    continue;

                // 1. Создаем параграф для элемента списка
                var listParagraph = doc.Content.Paragraphs.Add();

                // 2. Применяем форматированный текст
                content.ApplyToWord(doc, listParagraph.Range);

                // 3. Применяем форматирование списка
                if (IsOrdered)
                {
                    // Нумерованный список
                    listParagraph.Range.ListFormat.ApplyNumberDefault();
                }
                else
                {
                    // Маркированный список
                    listParagraph.Range.ListFormat.ApplyBulletDefault();
                }

                // 4. Добавляем перенос строки
                listParagraph.Range.InsertParagraphAfter();
                // Получаем последний параграф (который был создан через InsertParagraphAfter)
                var lastParagraphIndex = doc.Content.Paragraphs.Count;
                if (lastParagraphIndex > 0)
                {
                    var newParagraph = doc.Content.Paragraphs[lastParagraphIndex];
                    newParagraph.Range.ListFormat.RemoveNumbers();
                }
            }
        }
    }

    public class WordParagraph : IWordElement
    {
        public string ElementType => "Paragraph";
        public string StyleName { get; set; } = "Normal"; // для заголовков: Heading 1, Normal и т.д.
        public WdBuiltinStyle? BuiltinStyle { get; set; } = null;
        public WordFormattedText Content { get; set; } = new WordFormattedText();

        public WordParagraph(WdBuiltinStyle builtinStyle, WordFormattedText content)
        {
            BuiltinStyle = builtinStyle;
            StyleName = null; // Не используем строковое имя
            Content = content;
        }

        // Конструктор для обратной совместимости - принимает строковое имя стиля
        public WordParagraph(string styleName, WordFormattedText content)
        {
            StyleName = styleName;
            BuiltinStyle = null; // Не используем встроенную константу
            Content = content;
        }

        // вычисляемый уровень заголовка
        public int HeadingLevel
        {
            get
            {
                // Сначала проверяем встроенную константу
                if (BuiltinStyle.HasValue)
                {
                    WdBuiltinStyle style = BuiltinStyle.Value;
                    if (style == WdBuiltinStyle.wdStyleHeading1) return 1;
                    if (style == WdBuiltinStyle.wdStyleHeading2) return 2;
                    if (style == WdBuiltinStyle.wdStyleHeading3) return 3;
                    if (style == WdBuiltinStyle.wdStyleHeading4) return 4;
                    if (style == WdBuiltinStyle.wdStyleHeading5) return 5;
                    if (style == WdBuiltinStyle.wdStyleHeading6) return 6;
                }

                // Fallback: проверяем строковое имя (для обратной совместимости)
                if (!string.IsNullOrEmpty(StyleName))
                {
                    // Поддержка английских и русских названий
                    if (StyleName.StartsWith("Heading") || StyleName.StartsWith("Заголовок"))
                    {
                        string levelStr = StyleName
                            .Replace("Heading", "")
                            .Replace("Заголовок", "")
                            .Trim();
                        if (int.TryParse(levelStr, out int level))
                            return level;
                    }
                }

                return 0; // не заголовок
            }
        }

        public string ToMarkdown()
        {
            if (Content == null)
                return "";

            string contentMarkdown = Content.ToMarkdown();

            int headingLevel = HeadingLevel;
            if (headingLevel > 0)
            {
                // Преобразуем в заголовок Markdown
                string hashes = new string('#', headingLevel);
                return $"{hashes} {contentMarkdown}";
            }
            else
            {
                // Обычный параграф
                return contentMarkdown;
            }

        }

        public void ApplyToWord(Document doc)
        {
            if (doc == null)
                return;

            // 1. Создаем новый параграф
            var paragraph = doc.Content.Paragraphs.Add();
            // Отключаем форматирование списка, чтобы избежать наследования нумерации
            paragraph.Range.ListFormat.RemoveNumbers();

            // 2. Применяем стиль (Heading 1, Normal и т.д.)
            try
            {
                if (BuiltinStyle.HasValue)
                {
                    // Используем встроенную константу - автоматически локализуется
                    paragraph.Range.set_Style(BuiltinStyle.Value);
                }
                else if (!string.IsNullOrEmpty(StyleName))
                {
                    // Fallback на строковое имя (для обратной совместимости)
                    paragraph.Range.set_Style(StyleName);
                }
                else
                {
                    paragraph.Range.set_Style(WdBuiltinStyle.wdStyleNormal);
                }
            }
            catch
            {
                // Если стиль не существует, используем Normal
                paragraph.Range.set_Style(WdBuiltinStyle.wdStyleNormal);
            }

            // КРИТИЧЕСКИ ВАЖНО: Очищаем форматирование списка ПОСЛЕ применения стиля
            // потому что стили в Word могут иметь встроенную нумерацию
            paragraph.Range.ListFormat.RemoveNumbers();

            // Очищаем символ конца параграфа перед вставкой текста
            paragraph.Range.Text = "";

            // Очищаем форматирование списка после очистки текста
            paragraph.Range.ListFormat.RemoveNumbers();

            // 3. Применяем форматированный текст через Content
            if (Content != null && Content.Runs.Count > 0)
            {
                // Нужен метод ApplyToWord для WordFormattedText с Range
                Content.ApplyToWord(doc, paragraph.Range);
            }
            else
            {
                // Если нет форматирования, просто вставляем текст
                paragraph.Range.Text = Content?.ToMarkdown() ?? "";
            }

            // 4. Добавляем перенос строки после параграфа
            paragraph.Range.InsertParagraphAfter();

            // Очищаем форматирование списка у параграфа, созданного через InsertParagraphAfter
            // чтобы следующий элемент не наследовал нумерацию
            var lastParagraphIndex = doc.Content.Paragraphs.Count;
            if (lastParagraphIndex > 0)
            {
                var newParagraph = doc.Content.Paragraphs[lastParagraphIndex];
                newParagraph.Range.ListFormat.RemoveNumbers();
            }

        }

    }

    public class WordQuote : IWordElement
    {
        public string ElementType => "Quote";
        public string Text { get; set; }
        public WordFormattedText Content { get; set; }

        public WordQuote(string text, WordFormattedText content)
        {
            Text = text;
            Content = content;
        }
        public string ToMarkdown()
        {
            // Используем Content, если он есть, иначе Text
            string quoteText = "";
            if (Content != null)
            {
                quoteText = Content.ToMarkdown();
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                quoteText = Text;
            }

            if (string.IsNullOrEmpty(quoteText))
                return "";

            // Разбиваем на строки и добавляем префикс > для каждой строки
            var lines = quoteText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var sb = new StringBuilder();

            foreach (var line in lines)
            {
                sb.Append("> ");
                sb.AppendLine(line);
            }

            return sb.ToString().TrimEnd();
        }
        public void ApplyToWord(Document doc)
        {
            if (doc == null)
                return;

            // 1. Создаем параграф для цитаты
            var quoteParagraph = doc.Content.Paragraphs.Add();

            // Отключаем форматирование списка, чтобы избежать наследования нумерации
            quoteParagraph.Range.ListFormat.RemoveNumbers();

            // 2. Получаем текст цитаты
            //string quoteText = "";
            if (Content != null)
            {
                // Применяем форматированный текст
                Content.ApplyToWord(doc, quoteParagraph.Range);
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                quoteParagraph.Range.Text = Text;
            }

            // 3. Применяем стиль цитаты (если есть)
            try
            {
                quoteParagraph.Range.set_Style("Quote");
            }
            catch
            {
                // Если стиль "Quote" не существует, используем обычный стиль
                quoteParagraph.Range.set_Style("Normal");
                // Добавляем отступ для визуального выделения
                quoteParagraph.Range.ParagraphFormat.LeftIndent = 36; // 0.5 дюйма
            }

            // КРИТИЧЕСКИ ВАЖНО: Очищаем форматирование списка ПОСЛЕ применения стиля
            // потому что стили в Word могут иметь встроенную нумерацию
            quoteParagraph.Range.ListFormat.RemoveNumbers();

            // 4. Добавляем перенос строки
            quoteParagraph.Range.InsertParagraphAfter();

            // Очищаем форматирование списка у параграфа, созданного через InsertParagraphAfter
            // чтобы следующий элемент не наследовал нумерацию
            var lastParagraphIndex = doc.Content.Paragraphs.Count;
            if (lastParagraphIndex > 0)
            {
                var newParagraph = doc.Content.Paragraphs[lastParagraphIndex];
                newParagraph.Range.ListFormat.RemoveNumbers();
            }

        }

    }

    public class WordSubtitle : IWordElement
    {
        public string ElementType => "Subtitle";
        public string Text { get; set; }
        public WordFormattedText Content { get; set; }

        public WordSubtitle(string text, WordFormattedText content)
        {
            Text = text;
            Content = content;
        }
        public string ToMarkdown()
        {
            // Используем Content, если он есть, иначе Text
            string subtitleText = "";
            if (Content != null)
            {
                subtitleText = Content.ToMarkdown();
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                subtitleText = Text;
            }

            if (string.IsNullOrEmpty(subtitleText))
                return "";

            // Подзаголовок обычно представляется как заголовок уровня 2
            return $"## {subtitleText}";
        }
        public void ApplyToWord(Document doc)
        {

        }
    }

    public class WordTitle : IWordElement
    {
        public string ElementType => "Title";
        public string Text { get; set; }
        public WordFormattedText Content { get; set; }
        public int Level { get; set; }


        public WordTitle(string text, WordFormattedText content, int level)
        {
            Text = text;
            Content = content;
            Level = level;
        }

        private float GetFontSizeForLevel(int level)
        {
            switch (level)
            {
                case 1: return 24f;
                case 2: return 18f;
                case 3: return 14f;
                case 4: return 12f;
                case 5: return 11f;
                case 6: return 10f;
                default: return 12f; // значение по умолчанию
            }
        }

        public string ToMarkdown()
        {
            // Используем Content, если он есть, иначе Text
            string titleText = "";
            if (Content != null)
            {
                titleText = Content.ToMarkdown();
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                titleText = Text;
            }

            if (string.IsNullOrEmpty(titleText))
                return "";

            // Генерируем нужное количество символов # на основе Level
            string hashes = new string('#', Level);
            return $"{hashes} {titleText}";
        }

        public void ApplyToWord(Document doc)
        {
            if (doc == null)
            {
                System.Diagnostics.Debug.WriteLine("[WordTitle.ApplyToWord] doc равен null");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Начало применения заголовка уровня {Level}");

            // 1. Создаем параграф
            var paragraph = doc.Content.Paragraphs.Add();
            // Отключаем форматирование списка, чтобы избежать наследования нумерации
            paragraph.Range.ListFormat.RemoveNumbers();

            System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Параграф создан, Content: {(Content != null ? "не null" : "null")}, Runs.Count: {(Content?.Runs?.Count ?? 0)}");

            // 2. Получаем текст заголовка и создаем Range для применения стилей
            Range textRange = null;
            
            if (Content != null && Content.Runs.Count > 0)
            {
                string headingText = string.Join("", Content.Runs.Select(r => r?.Text ?? ""));
                System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Применение форматированного текста: '{headingText}'");

                // Вставляем текст
                Content.ApplyToWord(doc, paragraph.Range);

                // Используем paragraph.Range после вставки - он уже содержит только что вставленный текст
                textRange = paragraph.Range;
                System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] textRange создан: Start={textRange.Start}, End={textRange.End}, Text='{textRange.Text}'");
            }
            else if (!string.IsNullOrEmpty(Text))
            {
                System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Применение простого текста: '{Text}'");
                paragraph.Range.Text = Text;
                textRange = paragraph.Range;
            }
            else if (Content != null && Content.Runs.Count == 0)
            {
                // Если Content пустой, но не null, вставляем пустую строку
                System.Diagnostics.Debug.WriteLine("[WordTitle.ApplyToWord] ВНИМАНИЕ: Content не null, но Runs.Count == 0, вставляем пустую строку");
                paragraph.Range.Text = "";
                textRange = paragraph.Range;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[WordTitle.ApplyToWord] ВНИМАНИЕ: Content null и Text пустой!");
                textRange = paragraph.Range;
            }

            // 3. Устанавливаем размер шрифта на основе уровня
            float fontSize = GetFontSizeForLevel(Level);
            if (textRange != null)
            {
                textRange.Font.Size = fontSize;
                System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Размер шрифта установлен: {fontSize}");

                // 4. Делаем текст жирным
                textRange.Font.Bold = -1;

                // 5. Применяем стиль заголовка соответствующего уровня
                try
                {
                    WdBuiltinStyle headingStyle;
                    switch (Level)
                    {
                        case 1:
                            headingStyle = WdBuiltinStyle.wdStyleHeading1;
                            break;
                        case 2:
                            headingStyle = WdBuiltinStyle.wdStyleHeading2;
                            break;
                        case 3:
                            headingStyle = WdBuiltinStyle.wdStyleHeading3;
                            break;
                        case 4:
                            headingStyle = WdBuiltinStyle.wdStyleHeading4;
                            break;
                        case 5:
                            headingStyle = WdBuiltinStyle.wdStyleHeading5;
                            break;
                        case 6:
                            headingStyle = WdBuiltinStyle.wdStyleHeading6;
                            break;
                        default:
                            headingStyle = WdBuiltinStyle.wdStyleNormal;
                            System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] ВНИМАНИЕ: Неизвестный уровень заголовка {Level}, используется Normal");
                            break;
                    }
                    System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Применение стиля заголовка уровня {Level}");
                    textRange.set_Style(headingStyle);
                    System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Стиль применен успешно");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] ОШИБКА при применении стиля: {ex.Message}");
                    textRange.set_Style(WdBuiltinStyle.wdStyleNormal);
                }

                // КРИТИЧЕСКИ ВАЖНО: Очищаем форматирование списка ПОСЛЕ применения стиля
                // потому что стили заголовков в Word могут иметь встроенную нумерацию
                textRange.ListFormat.RemoveNumbers();
            }
            

            // 6. Добавляем перенос строки
            paragraph.Range.InsertParagraphAfter();

            // Очищаем форматирование списка у параграфа, созданного через InsertParagraphAfter
            // чтобы следующий элемент не наследовал нумерацию
            var lastParagraphIndex = doc.Content.Paragraphs.Count;
            if (lastParagraphIndex > 0)
            {
                var newParagraph = doc.Content.Paragraphs[lastParagraphIndex];
                newParagraph.Range.ListFormat.RemoveNumbers();
            }

            System.Diagnostics.Debug.WriteLine($"[WordTitle.ApplyToWord] Заголовок уровня {Level} успешно применен к Word");
        }
    }

    /// <summary>
    /// Класс для представления блока кода в Word документе.
    /// Реализует интерфейс IWordElement для преобразования между Markdown и Word форматами.
    /// </summary>
    public class WordCodeBlock : IWordElement
    {
        /// <summary>
        /// Возвращает тип элемента - "CodeBlock"
        /// </summary>
        public string ElementType => "CodeBlock";

        /// <summary>
        /// Текст кода, содержащийся в блоке
        /// </summary>
        public string Code { get; set; } = "";

        /// <summary>
        /// Язык программирования (опционально). Например: "csharp", "python", "javascript"
        /// Используется для подсветки синтаксиса в Markdown
        /// </summary>
        public string Language { get; set; } = "";

        /// <summary>
        /// Конструктор по умолчанию - создает пустой блок кода
        /// </summary>
        public WordCodeBlock()
        {
        }

        /// <summary>
        /// Конструктор с кодом
        /// </summary>
        /// <param name="code">Текст кода для блока</param>
        public WordCodeBlock(string code)
        {
            Code = code ?? ""; // Защита от null
        }

        /// <summary>
        /// Конструктор с кодом и языком программирования
        /// </summary>
        /// <param name="code">Текст кода для блока</param>
        /// <param name="language">Язык программирования (например, "csharp", "python")</param>
        public WordCodeBlock(string code, string language)
        {
            Code = code ?? "";
            Language = language ?? "";
        }

        /// <summary>
        /// Преобразует блок кода в строку Markdown.
        /// Формат: ```language\ncode\n```
        /// </summary>
        /// <returns>Строка Markdown, представляющая блок кода</returns>
        public string ToMarkdown()
        {
            // Если код пустой, возвращаем пустую строку
            if (string.IsNullOrEmpty(Code))
                return "";

            var sb = new StringBuilder();

            // Открывающие тройные кавычки
            sb.Append("```");

            // Добавляем язык, если он указан
            if (!string.IsNullOrEmpty(Language))
            {
                sb.Append(Language);
            }

            // Перенос строки после открывающих кавычек
            sb.AppendLine();

            // Добавляем сам код
            sb.Append(Code);

            // Закрывающие тройные кавычки на новой строке
            sb.AppendLine();
            sb.Append("```");

            return sb.ToString();
        }

        /// <summary>
        /// Применяет блок кода к Word документу.
        /// Создает параграф с моноширинным шрифтом и специальным форматированием.
        /// </summary>
        /// <param name="doc">Word документ, в который нужно вставить блок кода</param>
        public void ApplyToWord(Document doc)
        {
            // Проверка на null - защита от ошибок
            if (doc == null)
                return;

            // Если код пустой, ничего не делаем
            if (string.IsNullOrEmpty(Code))
                return;

            /*try
            {
                // 1. Создаем новый параграф для блока кода
                var codeParagraph = doc.Content.Paragraphs.Add();

                // 2. Сохраняем начальную позицию Range
                int startPosition = codeParagraph.Range.Start;

                // 3. Вставляем текст кода
                codeParagraph.Range.Text = Code;

                // 4. ВАЖНО: Обновляем Range после вставки текста
                // После установки Text, Range.End автоматически обновляется
                // Нужно получить актуальный Range с вставленным текстом
                int rangeStart = codeParagraph.Range.Start;
                int rangeEnd = codeParagraph.Range.End;
                
                // Создаём новый Range для работы с вставленным кодом
                // Исключаем символ конца параграфа (\r) из Range
                // Проверяем, что rangeEnd > rangeStart перед вычитанием
                int codeRangeEnd = rangeEnd > rangeStart ? rangeEnd - 1 : rangeEnd;
                Microsoft.Office.Interop.Word.Range codeRange = doc.Range(rangeStart, codeRangeEnd);

                // 5. Устанавливаем моноширинный шрифт (стандарт для кода)
                // Consolas - современный моноширинный шрифт
                codeRange.Font.Name = "Consolas";
                codeRange.Font.Size = 10;

                // 6. ПРИМЕНЯЕМ СТИЛЬ ПЕРЕД ПОДСВЕТКОЙ, чтобы не перезаписать цвета
                // Применяем стиль "Normal" для базового форматирования
                //codeRange.set_Style("Normal");

                // 7. Применяем визуальное форматирование (фон, отступы) ПЕРЕД подсветкой
                // Добавляем визуальное выделение через заливку
                codeRange.Shading.BackgroundPatternColor =
                    Microsoft.Office.Interop.Word.WdColor.wdColorGray25; // Светло-серый фон

                // Добавляем отступы для визуального выделения
                codeRange.ParagraphFormat.LeftIndent = 18; // 0.25 дюйма
                codeRange.ParagraphFormat.RightIndent = 18;

                // Добавляем отступ сверху и снизу
                codeRange.ParagraphFormat.SpaceBefore = 6;
                codeRange.ParagraphFormat.SpaceAfter = 6;

                // 8. ПРИМЕНЯЕМ ПОДСВЕТКУ СИНТАКСИСА ПОСЛЕ применения стиля и форматирования
                // Нормализуем язык (приводим к нижнему регистру)
                string normalizedLanguage = (Language ?? "").ToLower().Trim();
                if (string.IsNullOrEmpty(normalizedLanguage))
                {
                    normalizedLanguage = "python"; // Язык по умолчанию
                }

                // Применяем подсветку синтаксиса к вставленному коду
                // Это должно быть последним, чтобы цвета не были перезаписаны
                //SyntaxHighlighter.HighlightCodeBlock(codeRange, Code, normalizedLanguage);
                SyntaxHighlighterDiagnostics.DiagnoseTokenPositions(codeRange, Code, normalizedLanguage);

                // 9. Добавляем перенос строки после блока кода
                // Это создает визуальное разделение между кодом и следующим элементом
                codeParagraph.Range.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                // Обработка ошибок - выводим в отладочную консоль
                // В реальном приложении можно использовать логирование
                System.Diagnostics.Debug.WriteLine($"Ошибка при вставке блока кода: {ex.Message}");
            }*/

            try
            {
                // 1.Сохраняем начальную позицию
                int startPos = doc.Content.End - 1;

                // 2.Вставляем код через InsertAfter
                var insertRange = doc.Range(startPos);
                insertRange.InsertAfter(Code);

                // 3.Создаем Range для всего вставленного кода
                int endPos = doc.Content.End - 1;
                //Исключаем символ конца параграфа, если он есть
                if (endPos > startPos && doc.Range(endPos - 1, endPos).Text == "\r")
                {
                    endPos--;
                }
                var codeRange = doc.Range(startPos,endPos);

                // 4. Применяем форматирование к codeRange
                codeRange.Font.Name = "Consolas";
                codeRange.Font.Size = 10;
                codeRange.Shading.BackgroundPatternColorIndex = (WdColorIndex)WdColor.wdColorGray25;

                // 5. ПРИМЕНЯЕМ СТИЛЬ ПЕРЕД ПОДСВЕТКОЙ, чтобы не перезаписать цвета
                // Применяем стиль "Normal" для базового форматирования
                //codeRange.set_Style("Normal");

                // Добавляем отступы для визуального выделения
                codeRange.ParagraphFormat.LeftIndent = 18; // 0.25 дюйма
                codeRange.ParagraphFormat.RightIndent = 18;

                // Добавляем отступ сверху и снизу
                codeRange.ParagraphFormat.SpaceBefore = 6;
                codeRange.ParagraphFormat.SpaceAfter = 6;

                // 8. ПРИМЕНЯЕМ ПОДСВЕТКУ СИНТАКСИСА ПОСЛЕ применения стиля и форматирования
                // Нормализуем язык (приводим к нижнему регистру)
                string normalizedLanguage = (Language ?? "").ToLower().Trim();
                if (string.IsNullOrEmpty(normalizedLanguage))
                {
                    normalizedLanguage = "python"; // Язык по умолчанию
                }

                // Применяем подсветку синтаксиса к вставленному коду
                // Это должно быть последним, чтобы цвета не были перезаписаны
                // Теперь передаем только Range и язык - токены будут парситься из реального текста Word
                SyntaxHighlighter.HighlightCodeBlock(codeRange, normalizedLanguage);

                // 6. Добавляем пустой параграф после кода
                doc.Content.Paragraphs.Add();

            }
            catch (Exception ex) 
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка при вставке блока кода: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Класс для представления пустого параграфа (разделителя) в Word документе.
    /// Используется для обработки ThematicBreak (---) в Markdown.
    /// </summary>
    public class WordEmptyParagraph : IWordElement
    {
        /// <summary>
        /// Возвращает тип элемента - "EmptyParagraph"
        /// </summary>
        public string ElementType => "EmptyParagraph";

        /// <summary>
        /// Преобразует пустой параграф в строку Markdown (пустая строка)
        /// </summary>
        /// <returns>Пустая строка</returns>
        public string ToMarkdown()
        {
            return "";
        }

        /// <summary>
        /// Применяет пустой параграф к Word документу.
        /// Создает пустой параграф без текста и форматирования.
        /// </summary>
        /// <param name="doc">Word документ, в который нужно вставить пустой параграф</param>
        public void ApplyToWord(Document doc)
        {
            if (doc == null)
                return;

            // 1. Создаем новый параграф
            var paragraph = doc.Content.Paragraphs.Add();

            // 2. Отключаем форматирование списка, чтобы избежать наследования нумерации
            paragraph.Range.ListFormat.RemoveNumbers();

            // 3. Очищаем текст параграфа (делаем его пустым)
            paragraph.Range.Text = "";

            // 4. Применяем стиль Normal
            try
            {
                paragraph.Range.set_Style(WdBuiltinStyle.wdStyleNormal);
            }
            catch
            {
                // Игнорируем ошибки применения стиля
            }

            // 5. КРИТИЧЕСКИ ВАЖНО: Очищаем форматирование списка ПОСЛЕ применения стиля
            // потому что стили в Word могут иметь встроенную нумерацию
            paragraph.Range.ListFormat.RemoveNumbers();

            // 6. Добавляем перенос строки после пустого параграфа
            paragraph.Range.InsertParagraphAfter();

            // 7. Очищаем форматирование списка у параграфа, созданного через InsertParagraphAfter
            // чтобы следующий элемент не наследовал нумерацию
            var lastParagraphIndex = doc.Content.Paragraphs.Count;
            if (lastParagraphIndex > 0)
            {
                var newParagraph = doc.Content.Paragraphs[lastParagraphIndex];
                newParagraph.Range.ListFormat.RemoveNumbers();
            }
        }
    }

}