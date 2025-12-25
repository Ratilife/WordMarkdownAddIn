using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;



namespace WordMarkdownAddIn.Services
{
    public class WordToMarkdownService
    {
 
        /// <summary>
        /// Вспомогательный класс для хранения элемента документа вместе с его позицией.
        /// Используется для правильной сортировки элементов по порядку их появления в документе.
        /// </summary>
        private class ElementWithPosition
        {
            public IWordElement Element { get; set; }
            public int Position { get; set; }
        }

        private readonly Application _wordApp;
        private readonly Document _activeDoc;

        public WordToMarkdownService() 
        {
            _wordApp = Globals.ThisAddIn.Application;
            _activeDoc = _wordApp.ActiveDocument;

            if (_activeDoc == null)
                throw new System.Exception("Нет активного документа Word.");
        }

        public string ExtractDocumentContent()
        {
            return _activeDoc.Content.Text;
        }

        /// <summary>
        /// Извлекает структуру документа Word и преобразует её в список элементов IWordElement.
        /// Элементы извлекаются в порядке их появления в документе.
        /// </summary>
        /// <returns>Список элементов документа в порядке их появления.</returns>
        public List<IWordElement> ExtractDocumentStructure()
        {
           
            var elements = new List<ElementWithPosition>();
            var processedParagraphPositions = new HashSet<int>();

            // Извлекаем блоки кода ПЕРЕД параграфами (чтобы не обрабатывать их дважды)
            elements = ExtractCodeBlock(elements, processedParagraphPositions);
            
            // Извлекаем параграфы (теперь они сохраняют позицию при извлечении)
            // Пропускаем параграфы, которые уже обработаны как блоки кода
            elements = ExtractParagraphs(elements, processedParagraphPositions);
            
            // Извлекаем таблицы (теперь они сохраняют позицию при извлечении)
            elements = ExtractTables(elements);
            
            //  Просто сортируем по позиции и возвращаем только элементы
            // Больше не нужно вызывать GetElementStartPosition, так как позиция уже сохранена
            return elements
                .OrderBy(x => x.Position)
                .Select(x => x.Element)
                .ToList();
        }

  

        /// <summary>
        /// Преобразует структуру документа Word в строку Markdown.
        /// Извлекает все элементы документа, вызывает для каждого метод ToMarkdown()
        /// и объединяет результаты в одну строку Markdown.
        /// Результат готов для отображения в поле Markdown настройки.
        /// </summary>
        /// <returns>Строка с Markdown-представлением документа.</returns>
        public string ConvertToMarkdown()
        {
            try
            {
                // Извлекаем структуру документа
                var elements = ExtractDocumentStructure();
                
                if (elements == null || elements.Count == 0)
                    return "";

                var sb = new StringBuilder();
                
                // Проходим по всем элементам и преобразуем их в Markdown
                for (int i = 0; i < elements.Count; i++)
                {
                    var element = elements[i];
                    if (element == null)
                        continue;

                    string markdown = element.ToMarkdown();
                    
                    if (!string.IsNullOrEmpty(markdown))
                    {
                        sb.Append(markdown);
                        
                        // Добавляем переносы строк в зависимости от типа элемента
                        bool isLastElement = (i == elements.Count - 1);
                        
                        if (!isLastElement)
                        {
                            if (element is WordTable)
                            {
                                // Таблицы уже содержат переносы строк, добавляем одну пустую строку
                                sb.AppendLine();
                                sb.AppendLine();
                            }
                            else if (element is WordQuote)
                            {
                                // Цитаты уже содержат переносы строк, добавляем одну пустую строку
                                sb.AppendLine();
                                sb.AppendLine();
                            }
                            else if (element is WordParagraph para)
                            {
                                if (para.HeadingLevel > 0)
                                {
                                    // Заголовки - добавляем пустую строку после
                                    sb.AppendLine();
                                    sb.AppendLine();
                                }
                                else
                                {
                                    // Обычные параграфы - добавляем две пустые строки для разделения
                                    sb.AppendLine();
                                    sb.AppendLine();
                                }
                            }
                            else if (element is WordTitle || element is WordSubtitle)
                            {
                                // Заголовки и подзаголовки - добавляем пустую строку после
                                sb.AppendLine();
                                sb.AppendLine();
                            }
                            else if (element is WordListItem)
                            {
                                // Элементы списка - добавляем одну пустую строку
                                sb.AppendLine();
                            }
                            else if (element is WordCodeBlock)
                            {
                                // Блоки кода уже заканчиваются переносом строки в ToMarkdown()
                                // Добавляем одну пустую строку для разделения
                                sb.AppendLine();
                            }
                            else
                            {
                                // Другие элементы - добавляем две пустые строки
                                sb.AppendLine();
                                sb.AppendLine();
                            }
                        }
                    }
                }

                return sb.ToString().TrimEnd();
            }
            catch (Exception ex)
            {
                // В случае ошибки возвращаем пустую строку или можно логировать ошибку
                System.Diagnostics.Debug.WriteLine($"Ошибка при преобразовании в Markdown: {ex.Message}");
                return "";
            }
        }
        
        // Вспомогательный метод для извлечения форматирования
        private WordFormattedText ExtractFormattedContent(Range paragraphRange)
        {
            var formattedText = new WordFormattedText();
            var runs = formattedText.Runs;

            // Работаем с Characters для более точного анализа форматирования
            var chars = paragraphRange.Characters;

            if (chars.Count == 0) return formattedText; // Если параграф пустой

            // Начинаем с первого символа
            var firstChar = chars[1]; // Индексация с 1 в COM
            string currentText = firstChar.Text;
            var font = firstChar.Font; // Получаем шрифт первого символа
            var currentRun = new FormattedRun
            {
                Text = currentText,
                IsBold = font.Bold == 1, // Word использует -1 для true, 0 для false
                IsItalic = font.Italic == 1,
                IsUnderline = font.Underline != WdUnderline.wdUnderlineNone,
                IsStrikethrough = font.StrikeThrough != 0, // Зачеркивание
                IsSuperscript = font.Superscript != 0,     // Надстрочный индекс
                IsSubscript = font.Subscript != 0,         // Подстрочный индекс
                SmallCaps = font.SmallCaps != 0,           // Капитель
                AllCaps = font.AllCaps != 0,               // Все заглавные
            };

            // Проходим по остальным символам
            for (int i = 2; i <= chars.Count; i++)
            {
                var charRange = chars[i];
                var newFont = charRange.Font; // Получаем шрифт текущего символа

                // Проверяем, отличается ли форматирование от текущего фрагмента
                if (newFont.Bold == (currentRun.IsBold ? 1 : 0) &&
                    newFont.Italic == (currentRun.IsItalic ? 1 : 0) &&
                    (newFont.Underline != WdUnderline.wdUnderlineNone) == currentRun.IsUnderline &&
                    newFont.StrikeThrough == (currentRun.IsStrikethrough ? 1 : 0) && // Или != 0
                    newFont.Superscript == (currentRun.IsSuperscript ? 1 : 0) &&     // Или != 0
                    newFont.Subscript == (currentRun.IsSubscript ? 1 : 0) &&         // Или != 0
                    newFont.SmallCaps == (currentRun.SmallCaps ? 1 : 0) &&           // Или != 0
                    newFont.AllCaps == (currentRun.AllCaps ? 1 : 0) )                // Или != 0
                {
                    // Формат совпадает, добавляем символ к текущему фрагменту
                    currentRun.Text += charRange.Text;
                }
                else
                {
                    // Формат изменился, сохраняем текущий фрагмент
                    runs.Add(currentRun);

                    // Создаем новый фрагмент с новым форматированием
                    currentRun = new FormattedRun
                    {
                        Text = charRange.Text,
                        IsBold = font.Bold == 1,
                        IsItalic = font.Italic == 1,
                        IsUnderline = font.Underline != WdUnderline.wdUnderlineNone
                        // дописать
                    };
                }
            }

            // Не забываем добавить последний фрагмент
            if (!string.IsNullOrEmpty(currentRun.Text))
            {
                runs.Add(currentRun);
            }

            // Убираем символ конца параграфа из последнего фрагмента, если он там оказался
            if (runs.Count > 0 && runs[runs.Count - 1].Text.EndsWith("\r"))
            {
                runs[runs.Count - 1].Text = runs[runs.Count - 1].Text.TrimEnd('\r');
            }

            return formattedText;
        }

       
        // 1. Таблицы
        
        private List<ElementWithPosition> ExtractTables(List<ElementWithPosition> elements)
        {
            foreach (Table table in _activeDoc.Tables)
            {
                var tableData = new List<List<WordFormattedText>>();

                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    var rowData = new List<WordFormattedText>();
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        Range cellRange = table.Cell(i, j).Range;

                        
                        // Создаем копию Range без последнего символа (который обычно является маркером конца ячейки)
                        int cellStart = cellRange.Start;
                        int cellEnd = cellRange.End;

                        // Исключаем последний символ, если это маркер конца ячейки
                        if (cellEnd > cellStart)
                        {
                            cellEnd--; // Исключаем последний символ
                        }

                        // Создаем новый Range без маркера конца ячейки
                        Range cleanCellRange = _activeDoc.Range(cellStart, cellEnd);

                        // Извлекаем форматированный контент из очищенного Range
                        var formattedCellText = ExtractFormattedContent(cleanCellRange);
                        rowData.Add(formattedCellText);
                    }
                    tableData.Add(rowData);
                }
                
                //  Сохраняем элемент вместе с его позицией в документе
                var wordTable = new WordTable(tableData);
                elements.Add(new ElementWithPosition 
                { 
                    Element = wordTable, 
                    Position = table.Range.Start  // Сохраняем позицию начала таблицы
                });
            }

            return elements;
        }

        // 2. Параграфы
        
        private List<ElementWithPosition> ExtractParagraphs(List<ElementWithPosition> elements, HashSet<int> processedParagraphPositions)
        {
            // Обходим все параграфы
            foreach (Paragraph para in _activeDoc.Paragraphs)
            {
                // ПРОВЕРКА: Пропускаем параграфы, которые находятся внутри таблиц
                if (IsParagraphInTable(para))
                {
                    continue; // Пропускаем этот параграф
                }
                
                // ✅ ПРОВЕРКА: Пропускаем параграфы, которые уже обработаны как блоки кода
                int paragraphPosition = para.Range.Start;
                if (processedParagraphPositions.Contains(paragraphPosition))
                {
                    continue; // Пропускаем этот параграф, он уже в блоке кода
                }
                
                // Убираем символ конца параграфа
                string text = para.Range.Text.TrimEnd('\r', '\a');
                
                if (string.IsNullOrEmpty(text)) continue;
                
                //Определяем тип параграфа
                string styleName = para.get_Style().NameLocal;

                if (styleName.Contains("Heading") || styleName.Contains("Заголовок")) // Заголовок
                {
                    // Извлекаем форматированный контент
                    var content = ExtractFormattedContent(para.Range);
                    // Создаем WordParagraph с именем стиля и контентом.
                    // Уровень заголовка (HeadingLevel) будет вычислен автоматически при обращении к свойству.
                    // ✅ ИЗМЕНИТЬ: Сохраняем элемент вместе с позицией
                    elements.Add(new ElementWithPosition 
                    { 
                        Element = new WordParagraph(styleName, content), 
                        Position = paragraphPosition 
                    });
                    // Больше ничего делать не нужно, свойство HeadingLevel внутри WordParagraph сделает свою работу.
                }
                else if (styleName == "List Paragraph")
                {
                    // Проверяем тип списка через ListFormat
                    bool isOrdered = (para.Range.ListFormat.ListType != WdListType.wdListBullet);
                    // Это элемент списка
                    var content = ExtractFormattedContent(para.Range);
                    var contentList = new List<IWordElement> { content }; // Создаем список и добавляем один элемент (WordFormattedText является IWordElement)
                    // ✅ ИЗМЕНИТЬ: Сохраняем элемент вместе с позицией
                    elements.Add(new ElementWithPosition 
                    { 
                        Element = new WordListItem(contentList, isOrdered), 
                        Position = paragraphPosition 
                    });
                }
                else if (styleName == "Quote")
                {
                    // Это цитата
                    var content = ExtractFormattedContent(para.Range);
                    // ✅ ИЗМЕНИТЬ: Сохраняем элемент вместе с позицией
                    elements.Add(new ElementWithPosition 
                    { 
                        Element = new WordQuote(text, content), 
                        Position = paragraphPosition 
                    });
                }
                else if (styleName == "Subtitle")
                {
                    // Это подзаголовок
                    var content = ExtractFormattedContent(para.Range);
                    // ✅ ИЗМЕНИТЬ: Сохраняем элемент вместе с позицией
                    elements.Add(new ElementWithPosition 
                    { 
                        Element = new WordSubtitle(text, content), 
                        Position = paragraphPosition 
                    });
                }
                else if (styleName == "Title")
                {
                    // Это название документа
                    var content = ExtractFormattedContent(para.Range);
                    // ✅ ИЗМЕНИТЬ: Сохраняем элемент вместе с позицией
                    elements.Add(new ElementWithPosition 
                    {
                        Element = new WordTitle(text, content, 1),  // Title обычно соответствует уровню 1
                        Position = paragraphPosition 
                    });
                }
                else 
                {
                    if (!string.IsNullOrEmpty(text)) // Пример фильтрации
                    {
                        var content = ExtractFormattedContent(para.Range);
                        // ✅ ИЗМЕНИТЬ: Сохраняем элемент вместе с позицией
                        elements.Add(new ElementWithPosition 
                        { 
                            Element = new WordParagraph(para.get_Style().NameLocal, content), 
                            Position = paragraphPosition 
                        });
                    }
                }
            }
            return elements;

        }

        // Вспомогательный метод для проверки, находится ли параграф внутри таблицы
        private bool IsParagraphInTable(Paragraph para)
        {
            try
            {
                // Проверяем, пересекается ли Range параграфа с какой-либо таблицей
                foreach (Table table in _activeDoc.Tables)
                {
                    // Если начало или конец параграфа находится внутри Range таблицы
                    if (para.Range.Start >= table.Range.Start &&
                        para.Range.End <= table.Range.End)
                    {
                        return true; // Параграф находится внутри таблицы
                    }
                }
            }
            catch
            {
                // В случае ошибки считаем, что параграф не в таблице
            }

            return false; // Параграф не находится внутри таблицы
        }

        // 3. Гиперссылки
        private void ExtractHyperlinks()
        {

        }

        //  4. Изображения
        private void ExtractImages()
        {

        }

        // 5. Закладки
        private void ExtractBookmarks()
        {

        }

        // 6. Сноски
        private void ExtractFootnotes()
        {

        }
        //7.Блоки Кода 
        /// <summary>
        /// Извлекает блоки кода из документа Word.
        /// Блоки кода определяются по моноширинному шрифту (Consolas, Courier New, Courier и т.д.).
        /// Последовательные параграфы с моноширинным шрифтом группируются в один блок кода.
        /// </summary>
        /// <param name="elements">Список элементов для добавления найденных блоков кода</param>
        /// <param name="processedParagraphPositions">Множество позиций параграфов, которые уже обработаны (для пропуска в ExtractParagraphs)</param>
        /// <returns>Обновленный список элементов с добавленными блоками кода</returns>
        private List<ElementWithPosition> ExtractCodeBlock(List<ElementWithPosition> elements, HashSet<int> processedParagraphPositions)
        {
            // Список моноширинных шрифтов, которые используются для кода
            var monospaceFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Consolas",
                "Courier New",
                "Courier",
                "Lucida Console",
                "Monaco",
                "Menlo",
                "DejaVu Sans Mono",
                "Source Code Pro",
                "Fira Code"
            };

            // Подсчитываем общее количество параграфов для проверки размера блоков кода
            int totalParagraphs = 0;
            foreach (Paragraph para in _activeDoc.Paragraphs)
            {
                if (!IsParagraphInTable(para))
                {
                    totalParagraphs++;
                }
            }

            var codeBlockLines = new List<string>();
            int? codeBlockStartPosition = null;
            string detectedLanguage = "";
            int codeBlockParagraphCount = 0;
            var currentCodeBlockPositions = new List<int>(); // Отслеживаем позиции текущего блока кода

            foreach (Paragraph para in _activeDoc.Paragraphs)
            {
                // Пропускаем параграфы внутри таблиц
                if (IsParagraphInTable(para))
                {
                    // Если был начат блок кода, завершаем его
                    if (codeBlockStartPosition.HasValue)
                    {
                        FinishCodeBlock(elements, processedParagraphPositions, codeBlockLines, codeBlockStartPosition.Value, detectedLanguage);
                        codeBlockLines.Clear();
                        codeBlockStartPosition = null;
                        detectedLanguage = "";
                    }
                    continue;
                }

                // Получаем шрифт первого символа параграфа
                string fontName = "";
                bool hasGrayBackground = false;
                bool hasIndent = false;
                
                try
                {
                    if (para.Range.Characters.Count > 0)
                    {
                        fontName = para.Range.Characters[1].Font.Name;
                        
                        // Проверяем наличие серого фона (характерный признак блока кода)
                        try
                        {
                            var shadingColor = para.Range.Shading.BackgroundPatternColor;
                            // Серый фон обычно имеет значение wdColorGray25 или похожие
                            hasGrayBackground = (shadingColor == WdColor.wdColorGray25) ||
                                               (shadingColor == WdColor.wdColorGray15) ||
                                               (shadingColor == WdColor.wdColorGray10);
                        }
                        catch
                        {
                            // Если не удалось проверить фон, игнорируем
                        }
                        
                        // Проверяем наличие отступов (характерный признак блока кода)
                        try
                        {
                            hasIndent = para.Range.ParagraphFormat.LeftIndent > 0 ||
                                       para.Range.ParagraphFormat.RightIndent > 0;
                        }
                        catch
                        {
                            // Если не удалось проверить отступы, игнорируем
                        }
                    }
                }
                catch
                {
                    // Если не удалось получить шрифт, пропускаем
                }

                // Проверяем, является ли параграф частью блока кода
                // Блок кода должен иметь моноширинный шрифт И (серый фон ИЛИ отступы)
                // Это исключает случаи, когда весь документ имеет моноширинный шрифт
                bool isCodeParagraph = !string.IsNullOrEmpty(fontName) && 
                                      monospaceFonts.Contains(fontName) &&
                                      (hasGrayBackground || hasIndent);

                if (isCodeParagraph)
                {
                    // Убираем символ конца параграфа и пробелы в конце строки
                    string text = para.Range.Text.TrimEnd('\r', '\a', ' ', '\t');
                    
                    if (!string.IsNullOrEmpty(text))
                    {
                        // Если блок кода еще не начат, запоминаем начальную позицию
                        if (!codeBlockStartPosition.HasValue)
                        {
                            codeBlockStartPosition = para.Range.Start;
                            // Пытаемся определить язык по содержимому (опционально)
                            detectedLanguage = DetectLanguage(text);
                            codeBlockParagraphCount = 0;
                            currentCodeBlockPositions.Clear();
                        }
                        
                        // Добавляем строку кода
                        codeBlockLines.Add(text);
                        codeBlockParagraphCount++;
                        // Отмечаем параграф как обработанный и сохраняем позицию для возможного удаления
                        int paraPosition = para.Range.Start;
                        processedParagraphPositions.Add(paraPosition);
                        currentCodeBlockPositions.Add(paraPosition);
                    }
                }
                else
                {
                    // Если был начат блок кода, завершаем его
                    if (codeBlockStartPosition.HasValue)
                    {
                        // Проверяем, что блок кода не слишком большой (не более 50% документа)
                        // Это защита от случая, когда весь документ определяется как блок кода
                        double codeBlockPercentage = totalParagraphs > 0 ? 
                            (double)codeBlockParagraphCount / totalParagraphs * 100 : 0;
                        
                        if (codeBlockPercentage <= 50.0)
                        {
                            FinishCodeBlock(elements, processedParagraphPositions, codeBlockLines, codeBlockStartPosition.Value, detectedLanguage);
                        }
                        else
                        {
                            // Если блок кода слишком большой, не создаем его
                            // Удаляем позиции из processedParagraphPositions, чтобы эти параграфы были обработаны как обычные
                            foreach (int position in currentCodeBlockPositions)
                            {
                                processedParagraphPositions.Remove(position);
                            }
                        }
                        
                        codeBlockLines.Clear();
                        codeBlockStartPosition = null;
                        detectedLanguage = "";
                        codeBlockParagraphCount = 0;
                        currentCodeBlockPositions.Clear();
                    }
                }
            }

            // Завершаем последний блок кода, если он был начат
            if (codeBlockStartPosition.HasValue)
            {
                // Проверяем, что блок кода не слишком большой (не более 50% документа)
                double codeBlockPercentage = totalParagraphs > 0 ? 
                    (double)codeBlockParagraphCount / totalParagraphs * 100 : 0;
                
                if (codeBlockPercentage <= 50.0)
                {
                    FinishCodeBlock(elements, processedParagraphPositions, codeBlockLines, codeBlockStartPosition.Value, detectedLanguage);
                }
                else
                {
                    // Если блок кода слишком большой, не создаем его
                    // Удаляем позиции из processedParagraphPositions, чтобы эти параграфы были обработаны как обычные
                    foreach (int position in currentCodeBlockPositions)
                    {
                        processedParagraphPositions.Remove(position);
                    }
                }
            }

            return elements;
        }

        /// <summary>
        /// Завершает блок кода и добавляет его в список элементов
        /// </summary>
        private void FinishCodeBlock(List<ElementWithPosition> elements, HashSet<int> processedParagraphPositions, 
            List<string> codeLines, int startPosition, string language)
        {
            if (codeLines.Count == 0)
                return;

            // Нормализуем код: убираем лишние пустые строки в начале и конце
            var normalizedLines = NormalizeCodeLines(codeLines);

            if (normalizedLines.Count == 0)
                return;

            // Объединяем строки кода
            string codeText = string.Join("\n", normalizedLines);

            if (string.IsNullOrWhiteSpace(codeText))
                return;

            // Создаем WordCodeBlock
            var codeBlock = new WordCodeBlock(codeText, language);
            
            // Добавляем в список элементов
            elements.Add(new ElementWithPosition
            {
                Element = codeBlock,
                Position = startPosition
            });
        }

        /// <summary>
        /// Нормализует строки кода: убирает пустые строки в начале и конце, заменяет множественные пустые строки подряд на одну
        /// </summary>
        private List<string> NormalizeCodeLines(List<string> lines)
        {
            if (lines == null || lines.Count == 0)
                return new List<string>();

            var normalizedLines = new List<string>();
            bool lastWasEmpty = false;

            // Убираем пустые строки в начале
            int startIndex = 0;
            while (startIndex < lines.Count && string.IsNullOrWhiteSpace(lines[startIndex]))
            {
                startIndex++;
            }

            // Убираем пустые строки в конце
            int endIndex = lines.Count - 1;
            while (endIndex >= startIndex && string.IsNullOrWhiteSpace(lines[endIndex]))
            {
                endIndex--;
            }

            // Обрабатываем строки между началом и концом
            for (int i = startIndex; i <= endIndex; i++)
            {
                string line = lines[i];
                bool isEmpty = string.IsNullOrWhiteSpace(line);
                
                if (isEmpty)
                {
                    // Если предыдущая строка не была пустой, добавляем одну пустую строку
                    if (!lastWasEmpty)
                    {
                        normalizedLines.Add("");
                        lastWasEmpty = true;
                    }
                    // Иначе пропускаем (не добавляем множественные пустые строки подряд)
                }
                else
                {
                    // Удаляем пробелы в конце строки (но сохраняем отступы в начале)
                    string trimmedLine = line.TrimEnd();
                    normalizedLines.Add(trimmedLine);
                    lastWasEmpty = false;
                }
            }

            return normalizedLines;
        }

        /// <summary>
        /// Пытается определить язык программирования по содержимому кода
        /// </summary>
        private string DetectLanguage(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return "";

            string firstLine = code.Split('\n')[0].Trim();

            // Простые эвристики для определения языка
            if (firstLine.Contains("using ") && firstLine.Contains(";"))
                return "csharp";
            if (firstLine.Contains("def ") || firstLine.Contains("import ") || firstLine.Contains("from "))
                return "python";
            if (firstLine.Contains("function ") || firstLine.Contains("const ") || firstLine.Contains("let "))
                return "javascript";
            if (firstLine.Contains("public class") || firstLine.Contains("private "))
                return "java";
            if (firstLine.Contains("<?php") || firstLine.Contains("<?="))
                return "php";
            if (firstLine.Contains("SELECT") || firstLine.Contains("INSERT") || firstLine.Contains("UPDATE"))
                return "sql";
            if (firstLine.Contains("<!DOCTYPE") || firstLine.Contains("<html"))
                return "html";
            if (firstLine.Contains("package ") && firstLine.Contains("import "))
                return "go";

            return ""; // Язык не определен
        }
    }
}
