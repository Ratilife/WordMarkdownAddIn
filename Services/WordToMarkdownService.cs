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
            // ✅ ИЗМЕНИТЬ: Используем новый тип ElementWithPosition вместо IWordElement
            var elements = new List<ElementWithPosition>();

            // Извлекаем параграфы (теперь они сохраняют позицию при извлечении)
            elements = ExtractParagraphs(elements);
            
            // Извлекаем таблицы (теперь они сохраняют позицию при извлечении)
            elements = ExtractTables(elements);
            
            // ✅ ИЗМЕНИТЬ: Просто сортируем по позиции и возвращаем только элементы
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
        
        private List<ElementWithPosition> ExtractParagraphs(List<ElementWithPosition> elements )
        {
            // Обходим все параграфы
            foreach (Paragraph para in _activeDoc.Paragraphs)
            {
                // ПРОВЕРКА: Пропускаем параграфы, которые находятся внутри таблиц
                if (IsParagraphInTable(para))
                {
                    continue; // Пропускаем этот параграф
                }
                // Убираем символ конца параграфа
                string text = para.Range.Text.TrimEnd('\r', '\a');
                
                if (string.IsNullOrEmpty(text)) continue;
                
                // ✅ ДОБАВИТЬ: Сохраняем позицию параграфа для последующей сортировки
                int paragraphPosition = para.Range.Start;
                
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
                    var contentList = new List<WordFormattedText> { content }; // Создаем список и добавляем один элемент
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
        private void ExtractCodeBlock()
        {

        }
    }
}
